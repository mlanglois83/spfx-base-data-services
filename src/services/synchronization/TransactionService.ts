import { BaseDbService } from "../base/BaseDbService";
import { OfflineTransaction, SPFile } from "../../models/index";
import { assign } from "@microsoft/sp-lodash-subset";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { IAddOrUpdateResult } from "../../interfaces";

export class TransactionService extends BaseDbService<OfflineTransaction> {
    private transactionFileService: BaseDbService<SPFile>;

    constructor() {
        super( OfflineTransaction, "Transaction");
        this.transactionFileService = new BaseDbService<SPFile>(SPFile, "TransactionFiles");
    }

    /**
     * Add or update an item in DB and returns updated item
     * @param item Item to add or update
     */
    public async addOrUpdateItem(item: OfflineTransaction): Promise<IAddOrUpdateResult<OfflineTransaction>> {
        if (item.itemType === SPFile["name"]) {
            //create a file stored in a separate table
            let file: SPFile = assign(new SPFile(), item.itemData);
            item.itemData = new Date().getTime() + "_" + file.serverRelativeUrl;
            file.serverRelativeUrl = item.itemData;
            await this.transactionFileService.addOrUpdateItem(file);
        }
        let result = await super.addOrUpdateItem(item);
        return result;
    }

    public async deleteItem(item: OfflineTransaction): Promise<void> {
        if (item.itemType === SPFile["name"]) {
            let transaction = await super.getItemById(item.id);
            let file: SPFile = new SPFile();
            file.serverRelativeUrl = transaction.itemData;
            await this.transactionFileService.deleteItem(file);
        }
        await super.deleteItem(item);
    }


    /**
     * add items in table (ids updated)
     * @param newItems 
     */
    public async addOrUpdateItems(newItems: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        newItems = await Promise.all(newItems.map(async (item) => {
            if (item.itemType === SPFile["name"]) {
                //create a file stored in a separate table
                let file: SPFile = assign(new SPFile(), item.itemData);
                item.itemData = new Date().getTime() + "_" + file.serverRelativeUrl;
                file.serverRelativeUrl = item.itemData;
                await this.transactionFileService.addOrUpdateItem(file);
            }
            return item;
        }));
        newItems = await super.addOrUpdateItems(newItems);
        return newItems;
    }

    /**
     * Retrieve all items from db table
     */
    public async getAll(): Promise<Array<OfflineTransaction>> {
        let result = await super.getAll();
        result = await Promise.all(result.map(async (item) => {
            if (item.itemType === SPFile["name"]) {
                let file = await this.transactionFileService.getItemById(item.itemData);
                if (file) {
                    file.serverRelativeUrl = file.serverRelativeUrl.replace(/^\d+_(.*)$/g, "$1");
                    item.itemData = assign({}, file);
                }
            }
            return item;
        }));
        return result;
    }


    /**
     * Clear table
     */
    public async clear(): Promise<void> {
        await this.transactionFileService.clear();
        await super.clear();
    }
}