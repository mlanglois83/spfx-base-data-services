
import { BaseService } from "../base/BaseService";
import { BaseDbService } from "../base/BaseDbService";
import { BaseItem, OfflineTransaction, SPFile } from "../../models/index";
import { TransactionType, Constants } from "../../constants/index";
import { assign } from "@microsoft/sp-lodash-subset";
import { IItemSynchronized, ISynchronizationEnded } from "../../interfaces/index";
import { TransactionService } from "./TransactionService";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { ServiceFactory } from "../ServiceFactory";


export class SynchronizationService extends BaseService {
    private transactionService: BaseDbService<OfflineTransaction>;

    private static itemSynchroCallbacks = {};
    private static synchroCallbacks = {};

    /**
     * Registers a function called when an item was synchronized
     * @param key - unique key for callback
     * @param callback - callback function called when an item was synchronized
     */
    public static registerItemSynchronizedCallback(key: string, callback: (synchro: IItemSynchronized) => void): void {
        SynchronizationService.itemSynchroCallbacks[key] = callback;
    }
    /**
     * Unregisters a function associated with item synchronisation
     * @param key - unique callback key
     */
    public static unregisterItemSynchronizedCallback(key: string): void {
        if (SynchronizationService.itemSynchroCallbacks[key]) {
            delete (SynchronizationService.itemSynchroCallbacks[key]);
        }
    }
    /**
     * Registers a function called when synchronization has ended
     * @param key - unique key for callback
     * @param callback - callback function called when synchronization has ended
     */
    public static registerSynchronizationCallback(key: string, callback: (synchroResult: ISynchronizationEnded) => void): void {
        SynchronizationService.synchroCallbacks[key] = callback;
    }
    /**
     * Unregister a function registered for synchronisation end
     * @param key - unique callback key
     */
    public static unregisterSynchronizationCallback(key: string): void {
        if (SynchronizationService.synchroCallbacks[key]) {
            delete (SynchronizationService.synchroCallbacks[key]);
        }
    }

    private static emitItemSynchronized(synchro: IItemSynchronized): void {
        for (const key in SynchronizationService.itemSynchroCallbacks) {
            if (SynchronizationService.itemSynchroCallbacks.hasOwnProperty(key)) {
                const callback = SynchronizationService.itemSynchroCallbacks[key];
                callback(synchro);
            }
        }
    }
    private static emitSynchronizationEnded(synchro: ISynchronizationEnded): void {
        for (const key in SynchronizationService.synchroCallbacks) {
            if (SynchronizationService.synchroCallbacks.hasOwnProperty(key)) {
                const callback = SynchronizationService.synchroCallbacks[key];
                callback(synchro);
            }
        }
    }

    constructor() {
        super();
        this.transactionService = new TransactionService();
    }


    public async run(): Promise<Array<string>> {
        const errors = [];
        //read transaction table
        const transactions = await this.transactionService.getAll();
        for (let index = 0; index < transactions.length; index++) {
            const transaction = transactions[index];
            // get associated type & service
            const dataService = ServiceFactory.getServiceByModelName(transaction.itemType);
            // init service for tardive links
            await dataService.Init();
            // transform item to destination type
            const item: BaseItem = assign(ServiceFactory.getItemByName(transaction.itemType), transaction.itemData);
            switch (transaction.title) {
                case TransactionType.AddOrUpdate:
                    const oldId = item.id;
                    const isAdd = typeof (oldId) === "number" && oldId < 0;
                    const tmp = await dataService.mapItems([item]);
                    const dbItem = tmp.shift();
                    const updatedItem = await dataService.addOrUpdateItem(dbItem);

                    // handle id and version changed
                    if (isAdd && !updatedItem.error) {

                        let nextTransactions: Array<OfflineTransaction> = [];
                        // next transactions on this item
                        if (index < transactions.length - 1) {
                            nextTransactions = await Promise.all(transactions.slice(index + 1).map(async (updatedTr) => {

                                if (updatedTr.itemType === transaction.itemType && (updatedTr.itemData as BaseItem).id === oldId) {
                                    (updatedTr.itemData as BaseItem).id = updatedItem.id;
                                    (updatedTr.itemData as BaseItem).version = updatedItem.version;

                                    const identifiers = dataService.Identifier;
                                    if (identifiers) {
                                        for (const identifier of identifiers) {

                                            (updatedTr.itemData as BaseItem)[identifier] = updatedItem[identifier];
                                        }
                                    }

                                    await this.transactionService.addOrUpdateItem(updatedTr);
                                }
                                return updatedTr;
                            }));
                        }
                        if (dataService.updateLinkedTransactions) {
                            nextTransactions = await dataService.updateLinkedTransactions(oldId, updatedItem.id, nextTransactions);
                        }
                        if (index < transactions.length - 1) {
                            transactions.splice(index + 1, transactions.length - index - 1, ...nextTransactions);
                        }

                    }
                    // update version on next transactions (avoid errors)
                    else if (!updatedItem.error) {
                        let nextTransactions: Array<OfflineTransaction> = [];
                        // next transactions on this item
                        if (index < transactions.length - 1) {
                            nextTransactions = await Promise.all(transactions.slice(index + 1).map(async (updatedTr) => {
                                if (updatedTr.itemType === transaction.itemType &&
                                    (updatedTr.itemData as BaseItem).id === item.id) {
                                    (updatedTr.itemData as BaseItem).version = updatedItem.version;
                                    await this.transactionService.addOrUpdateItem(updatedTr);
                                }
                                return updatedTr;
                            }));
                        }
                        if (index < transactions.length - 1) {
                            transactions.splice(index + 1, transactions.length - index - 1, ...nextTransactions);
                        }
                    }
                    if (updatedItem.error) {
                        errors.push(this.formatError(transaction, updatedItem.error.message));
                        if (updatedItem.error.name === Constants.Errors.ItemVersionConfict) {
                            await this.transactionService.deleteItem(transaction);
                        }
                    }
                    else {
                        await this.transactionService.deleteItem(transaction);
                    }
                    SynchronizationService.emitItemSynchronized({ item: updatedItem, oldId: (isAdd ? oldId : undefined), operation: TransactionType.AddOrUpdate });
                    break;
                case TransactionType.Delete:
                    try {
                        await dataService.deleteItem(item);
                        await this.transactionService.deleteItem(transaction);
                        SynchronizationService.emitItemSynchronized({ item: item, operation: TransactionType.Delete });
                    } catch (error) {
                        errors.push(this.formatError(transaction, error.message));
                    }
                    break;
            }
        }

        SynchronizationService.emitSynchronizationEnded({ errors: errors });
        //return errors list
        return errors;
    }

    private formatError(transaction: OfflineTransaction, message: string): string {
        let operationLabel: string;
        const item = assign(ServiceFactory.getItemByName(transaction.itemType), transaction.itemData);
        switch (transaction.title) {
            case TransactionType.AddOrUpdate:
                if (item instanceof SPFile) {
                    operationLabel = ServicesConfiguration.configuration.translations.UploadLabel;
                }
                else if (item.id < 0) {
                    operationLabel = ServicesConfiguration.configuration.translations.AddLabel;
                }
                else {
                    operationLabel = ServicesConfiguration.configuration.translations.UpdateLabel;
                }
                break;
            case TransactionType.Delete:
                operationLabel = ServicesConfiguration.configuration.translations.DeleteLabel;
                break;
            default: break;
        }
        const itemTypeLabel = ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType] ? ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType] : transaction.itemType;
        return Text.format(ServicesConfiguration.configuration.translations.SynchronisationErrorFormat, itemTypeLabel, operationLabel, item.title, item.id, message);
    }

}