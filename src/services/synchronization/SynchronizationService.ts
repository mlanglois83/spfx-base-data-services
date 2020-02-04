
import { BaseService } from "../base/BaseService";
import { BaseDbService } from "../base/BaseDbService";
import { OfflineTransaction, SPFile } from "../../models/index";
import { TransactionType, Constants } from "../../constants/index";
import { BaseServiceFactory } from "../base/BaseServiceFactory";
import { assign } from "@microsoft/sp-lodash-subset";
import { IBaseItem } from "../../interfaces/index";
import { TransactionService } from "./TransactionService";
import { Text } from "@microsoft/sp-core-library";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";


export class SynchronizationService extends BaseService {
    private transactionService: BaseDbService<OfflineTransaction>;


    constructor() {
        super();
        this.transactionService = new TransactionService();

    }


    public async run(): Promise<Array<string>> {
        let errors = [];
        //read transaction table
        let transactions = await this.transactionService.getAll();
        for (let index = 0; index < transactions.length; index++) {
            const transaction = transactions[index];
            // get associated type & service
            let itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
            let dataService = ServicesConfiguration.configuration.serviceFactory.create(transaction.itemType);
            // init service for tardive links
            await dataService.Init();
            // transform item to destination type
            let item = assign(new itemType(), transaction.itemData);
            switch (transaction.title) {
                case TransactionType.AddOrUpdate:
                    const oldId = item.id;
                    const isAdd = typeof (oldId) === "number" && oldId < 0;       
                    let dbItem = await dataService.mapItem(item);
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
                        if (dataService.updateLinkedTransactions) {
                            nextTransactions = await dataService.updateLinkedTransactions(oldId, updatedItem.item.id, nextTransactions);
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
                        // TODO
                    }
                    else {
                        await this.transactionService.deleteItem(transaction);
                        
                        // TODO: on item updated (item)
                    }
                    break;
                case TransactionType.Delete:
                    try {
                        await dataService.deleteItem(item);
                        await this.transactionService.deleteItem(transaction);
                        // TODO: OnItemDeleted (type + id)
                    } catch (error) {
                        errors.push(this.formatError(transaction, error.message));
                    }
                    break;
            }
        }        
        //return errors list
        return errors;
    }

    private formatError(transaction: OfflineTransaction, message: string) {
        let operationLabel: string;
        let itemTypeLabel :string;
        let itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
        let item = assign(new itemType(), transaction.itemData);
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
        itemTypeLabel = ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType] ? ServicesConfiguration.configuration.translations.typeTranslations[transaction.itemType]: transaction.itemType;
        return Text.format(ServicesConfiguration.configuration.translations.SynchronisationErrorFormat, itemTypeLabel, operationLabel, item.title, item.id, message);
    }

}