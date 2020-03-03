import { assign } from "@microsoft/sp-lodash-subset";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { OfflineTransaction, SPFile } from "../../models/index";
import { BaseDbService } from "../base/BaseDbService";

export class TransactionService extends BaseDbService<OfflineTransaction> {
    private transactionFileService: BaseDbService<SPFile>;

    constructor() {
        super( OfflineTransaction, "Transaction");
        this.transactionFileService = new BaseDbService<SPFile>(SPFile, "TransactionFiles");
    }



    /**
     * Add or update an item in DB and returns updated item
     * @param item - item to add or update
     */
    public async addOrUpdateItem(item: OfflineTransaction): Promise<OfflineTransaction> {
        let result: OfflineTransaction = null;
        if (this.isFile(item.itemType)) {
            // if existing transaction, remove with associated files
            const existing = await this.getItemById(item.id);
            if(existing) {
                await this.deleteItem(existing);
            }
            //create a file stored in a separate table
            const file: SPFile = assign(new SPFile(), item.itemData);
            const baseUrl = file.serverRelativeUrl;
            item.itemData = new Date().getTime() + "_" + file.serverRelativeUrl;
            file.serverRelativeUrl = item.itemData;
            await this.transactionFileService.addOrUpdateItem(file);            
            result = await super.addOrUpdateItem(item);
            // reassign values for result
            file.serverRelativeUrl = baseUrl;
            result.itemData = assign({}, file);
        }
        else {
            result = await super.addOrUpdateItem(item);
        }
        return result;
    }

    public async deleteItem(item: OfflineTransaction): Promise<void> {
        if (this.isFile(item.itemType)) {
            const transaction = await super.getItemById(item.id);
            const file: SPFile = new SPFile();
            file.serverRelativeUrl = transaction.itemData;
            await this.transactionFileService.deleteItem(file);
        }
        await super.deleteItem(item);
    }

    /**
     * add items in table (ids updated)
     * @param newItems  - items to add or update
     */
    public async addOrUpdateItems(newItems: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        const updateResults = Promise.all(newItems.map(async (item) => {
            const result = await this.addOrUpdateItem(item);
            return result;
        }));
        return updateResults;
    }

    /**
     * Retrieve all items from db table
     */
    public async getAll(): Promise<Array<OfflineTransaction>> {
        let result = await super.getAll();
        result = await Promise.all(result.map(async (item) => {            
            if (this.isFile(item.itemType)) {
                const file = await this.transactionFileService.getItemById(item.itemData);
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
     * Get a transaction given its id
     * @param id - transaction id
     */
    public async getItemById(id: number): Promise<OfflineTransaction> {
        const result = await super.getItemById(id);
        if (result && result.itemType === SPFile["name"]) {
            const file = await this.transactionFileService.getItemById(result.itemData);
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

    private isFile(itemTypeName: string): boolean {
        const itemType = ServicesConfiguration.configuration.serviceFactory.getItemTypeByName(itemTypeName);
        const instance = new itemType();
        return (instance instanceof SPFile);
    }

}