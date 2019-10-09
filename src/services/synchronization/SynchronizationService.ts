
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
        await Promise.all(transactions.map((transaction, index) => {
            return new Promise<void>(async(resolve, reject) => {
                // get associated type & service
                let itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(transaction.itemType);
                let dataService = ServicesConfiguration.configuration.serviceFactory.create(transaction.serviceName);
                // transform item to destination type
                let item = assign(new itemType(), transaction.itemData);
                switch (transaction.title) {
                    case TransactionType.AddOrUpdate:
                        const oldId = item.id;
                        const isAdd = typeof (oldId) === "number" && oldId < 0;                        
                        const updatedItem = await dataService.addOrUpdateItem(item);
                        // handle id and version changed
                        if (isAdd && !updatedItem.error) {
                            // next transactions on this item
                            if (index < transactions.length - 1) {
                                transactions.slice(index).filter((t) => {
                                    return t.itemType === transaction.itemType &&
                                        t.serviceName === transaction.serviceName &&
                                        (t.itemData as IBaseItem).id === oldId;
                                }).forEach(async (updatedTr) => {
                                    (updatedTr.itemData as IBaseItem).id = updatedItem.item.id;
                                    (updatedTr.itemData as IBaseItem).version = updatedItem.item.version;
                                    await this.transactionService.addOrUpdateItem(updatedTr);
                                });
                            }
                            // other update for linked content
                            if (dataService.updateLinkedItems) {
                                dataService.updateLinkedItems(oldId, updatedItem.item.id);
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
                        resolve();
                        break;
                    case TransactionType.Delete:
                        try {
                            await dataService.deleteItem(item);
                            await this.transactionService.deleteItem(transaction);
                            resolve();
                        } catch (error) {
                            errors.push(this.formatError(transaction, error.message));
                            resolve();
                        }
                        break;
                }
            });            
        }));
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