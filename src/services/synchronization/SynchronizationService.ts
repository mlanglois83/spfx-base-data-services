
import { BaseService } from "../base/BaseService";
import { BaseDbService } from "../base/BaseDbService";
import { OfflineTransaction, SPFile } from "../../models/index";
import { TransactionType, Constants } from "../../constants/index";
import { assign } from "@microsoft/sp-lodash-subset";
import { IBaseItem, IItemSynchronized, ISynchronizationEnded } from "../../interfaces/index";
import { TransactionService } from "./TransactionService";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";


export class SynchronizationService extends BaseService {
    private transactionService: BaseDbService<OfflineTransaction>;

    private static itemSynchroCallbacks = {};
    private static synchroCallbacks = {};

    /**
     * Registers a function called when an item was synchronized
     * @param key Unique key for callback
     * @param callback Callbackfunction called when an item was synchronized
     */
    public static registerItemSynchronizedCallback(key: string, callback: (synchro: IItemSynchronized) => void): void {
        SynchronizationService.itemSynchroCallbacks[key] = callback;
    }
    /**
     * Unregisters a function associated with item synchronisation
     * @param key Unique callback key
     */
    public static unregisterItemSynchronizedCallback(key: string): void {
        if(SynchronizationService.itemSynchroCallbacks[key]) {
            delete(SynchronizationService.itemSynchroCallbacks[key]);
        }
    }
    /**
     * Registers a function called when synchronization has ended
     * @param key Unique key for callback
     * @param callback Callbackfunction called when  synchronization has ended
     */
    public static registerSynchronizationCallback(key: string, callback: (synchroResult: ISynchronizationEnded) => void): void {
        SynchronizationService.synchroCallbacks[key] = callback;
    }
    /**
     * Unregister a function registered for synchronisation end
     * @param key Unique callback key
     */
    public static unregisterSynchronizationCallback(key: string): void {
        if(SynchronizationService.synchroCallbacks[key]) {
            delete(SynchronizationService.synchroCallbacks[key]);
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
            const itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
            const dataService = ServicesConfiguration.configuration.serviceFactory.create(transaction.itemType);
            // init service for tardive links
            await dataService.Init();
            // transform item to destination type
            const item = assign(new itemType(), transaction.itemData);
            switch (transaction.title) {
                case TransactionType.AddOrUpdate:
                    const oldId = item.id;
                    const isAdd = typeof (oldId) === "number" && oldId < 0;       
                    const dbItem = await dataService.mapItem(item);
                    const updatedItem = await dataService.addOrUpdateItem(dbItem);

                    // handle id and version changed
                    if (isAdd && !updatedItem.error) {
                        
                        let nextTransactions: Array<OfflineTransaction> = [];
                        // next transactions on this item
                        if (index < transactions.length - 1) {
                            nextTransactions = await Promise.all(transactions.slice(index + 1).map(async (updatedTr) => {
                                if(updatedTr.itemType === transaction.itemType &&
                                (updatedTr.itemData as IBaseItem).id === oldId) {
                                    (updatedTr.itemData as IBaseItem).id = updatedItem.item.id;
                                    (updatedTr.itemData as IBaseItem).version = updatedItem.item.version;
                                    await this.transactionService.addOrUpdateItem(updatedTr);
                                }
                                return updatedTr;                            
                            }));
                        }
                        if (dataService.updateLinkedTransactions) {
                            nextTransactions = await dataService.updateLinkedTransactions(oldId, updatedItem.item.id, nextTransactions);
                        }
                        if(index < transactions.length - 1) {
                            transactions.splice(index + 1, transactions.length - index - 1, ...nextTransactions);
                        }

                    }
                    // update version on next transactions (avoid errors)
                    else if(!updatedItem.error) {
                        let nextTransactions: Array<OfflineTransaction> = [];
                        // next transactions on this item
                        if (index < transactions.length - 1) {
                            nextTransactions = await Promise.all(transactions.slice(index + 1).map(async (updatedTr) => {
                                if(updatedTr.itemType === transaction.itemType &&
                                (updatedTr.itemData as IBaseItem).id === item.id) {
                                    (updatedTr.itemData as IBaseItem).version = updatedItem.item.version;
                                    await this.transactionService.addOrUpdateItem(updatedTr);
                                }
                                return updatedTr;                            
                            }));
                        }
                        if(index < transactions.length - 1) {
                            transactions.splice(index + 1, transactions.length - index - 1, ...nextTransactions);
                        }
                    }
                    if(updatedItem.error) {
                        errors.push(this.formatError(transaction, updatedItem.error.message));
                        if(updatedItem.error.name === Constants.Errors.ItemVersionConfict){
                            await this.transactionService.deleteItem(transaction);
                        }
                    }
                    else {
                        await this.transactionService.deleteItem(transaction);
                    }
                    SynchronizationService.emitItemSynchronized({item: updatedItem.item, oldId: (isAdd ? oldId: undefined), operation: TransactionType.AddOrUpdate});
                    break;
                case TransactionType.Delete:
                    try {
                        await dataService.deleteItem(item);
                        await this.transactionService.deleteItem(transaction);
                        SynchronizationService.emitItemSynchronized({item: item,operation: TransactionType.Delete});
                    } catch (error) {
                        errors.push(this.formatError(transaction, error.message));
                    }
                    break;
            }
        }        
        
        SynchronizationService.emitSynchronizationEnded({errors: errors});
        //return errors list
        return errors;
    }

    private formatError(transaction: OfflineTransaction, message: string): string {
        let operationLabel: string;
        const itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
        const item = assign(new itemType(), transaction.itemData);
        switch (transaction.title) {
            case TransactionType.AddOrUpdate:
                if(item instanceof SPFile) {
                    operationLabel = ServicesConfiguration.configuration.translations.UploadLabel;
                }
                else if(item.id < 0) {
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
        const itemTypeLabel = ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType] ? ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType]: transaction.itemType;
        return Text.format(ServicesConfiguration.configuration.translations.SynchronisationErrorFormat, itemTypeLabel, operationLabel, item.title, item.id, message);
    }

}