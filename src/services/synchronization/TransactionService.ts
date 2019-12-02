import { BaseDbService } from "../base/BaseDbService";
import { OfflineTransaction, SPFile } from "../../models/index";
import { assign, update } from "@microsoft/sp-lodash-subset";
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
        let result : IAddOrUpdateResult<OfflineTransaction> = null;
        if (item.itemType === SPFile["name"]) {
            // if existing transaction, remove with associated files
            let existing = await this.getById(item.id);
            if(existing) {
                await this.deleteItem(existing);
            }
            //create a file stored in a separate table
            let file: SPFile = assign(new SPFile(), item.itemData);
            let baseUrl = file.serverRelativeUrl;
            item.itemData = new Date().getTime() + "_" + file.serverRelativeUrl;
            file.serverRelativeUrl = item.itemData;
            await this.transactionFileService.addOrUpdateItem(file);            
            result = await super.addOrUpdateItem(item);
            // reassign values for result
            file.serverRelativeUrl = baseUrl;
            result.item.itemData = assign({}, file);
        }
        else {
            result = await super.addOrUpdateItem(item);
        }
        return result;
    }

    public async deleteItem(item: OfflineTransaction): Promise<void> {
        if (item.itemType === SPFile["name"]) {
            let transaction = await super.getById(item.id);
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
        let updateResults = Promise.all(newItems.map(async (item) => {
            let result = await this.addOrUpdateItem(item);
            return result.item;
        }));
        return updateResults;
    }

    /**
     * Retrieve all items from db table
     */
    public async getAll(): Promise<Array<OfflineTransaction>> {
        let result = await super.getAll();
        result = await Promise.all(result.map(async (item) => {
            if (item.itemType === SPFile["name"]) {
                let file = await this.transactionFileService.getById(item.itemData);
                if (file) {
                    file.serverRelativeUrl = file.serverRelativeUrl.replace(/^\d+_(.*)$/g, "$1");
                    item.itemData = assign({}, file);
                }
            }
            return item;
        }));
        return result;
    }

    public async getById(id: number): Promise<OfflineTransaction> {
        let result = await super.getById(id);
        if (result && result.itemType === SPFile["name"]) {
            let file = await this.transactionFileService.getById(result.itemData);
            if (file) {
                file.serverRelativeUrl = file.serverRelativeUrl.replace(/^\d+_(.*)$/g, "$1");
                result.itemData = assign({}, file);
            }
        }
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