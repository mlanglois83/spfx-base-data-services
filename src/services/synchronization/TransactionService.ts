import { assign } from "@microsoft/sp-lodash-subset";
import { ServiceFactory } from "../ServiceFactory";
import { BaseFile, OfflineTransaction } from "../../models/index";
import { BaseDbService } from "../base/BaseDbService";

export class TransactionService extends BaseDbService<OfflineTransaction> {
    private transactionFileService: BaseDbService<BaseFile>;

    constructor() {
        super( OfflineTransaction, "OfflineTransaction");
        this.transactionFileService = new BaseDbService<BaseFile>(BaseFile, "OfflineTransactionFiles");
    }



    /**
     * Add or update an item in DB and returns updated item
     * @param item - item to add or update
     */
    public async addOrUpdateItem(item: OfflineTransaction): Promise<OfflineTransaction> {
        let result: OfflineTransaction = null;
        if (this.isFile(item.itemType)) {
            // if existing transaction, remove with associated files
            const existing = await super.getItemById(item.id);
            if(existing) {
                await this.deleteItem(existing);
            }
            //create a file stored in a separate table
            const file: BaseFile = assign(new BaseFile(), item.itemData);
            const baseUrl = file.id;
            item.itemData = new Date().getTime() + "_" + file.id;
            file.id = item.itemData;
            await this.transactionFileService.addOrUpdateItem(file);            
            result = await super.addOrUpdateItem(item);
            // reassign values for result
            file.id = baseUrl;
            result.itemData = assign({}, file);
        }
        else {
            result = await super.addOrUpdateItem(item);
        }
        return result;
    }

    public async deleteItem(item: OfflineTransaction): Promise<OfflineTransaction> {
        if (this.isFile(item.itemType)) {
            const transaction = await super.getItemById(item.id);
            const file: BaseFile = new BaseFile();
            file.id = transaction.itemData;
            await this.transactionFileService.deleteItem(file);
        }
        return super.deleteItem(item);
    }
    public async deleteItems(items: Array<OfflineTransaction>): Promise<Array<OfflineTransaction>> {
        const files = [];
        // remove files
        await Promise.all(items.filter(item => this.isFile(item.itemType)).map(async(item) => {
            if(this.isFile(item.itemType)) {
                const transaction = await super.getItemById(item.id);
                const file: BaseFile = new BaseFile();
                file.id = transaction.itemData;
                files.push(file);
            }
        }));
        if(files.length > 0) {
            await this.transactionFileService.deleteItems(files);
        }
        super.deleteItems(items);
        return items;
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
                    const id = file.id.toString().replace(/^\d+_(.*)$/g, "$1");
                    const temp = ServiceFactory.getItemByName(item.itemType);
                    if(typeof(temp.id) === "number") {
                        file.id = parseInt(id);
                    }
                    else {
                        file.id = id;
                    }
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
        if (this.isFile(result.itemType)) {
            const file = await this.transactionFileService.getItemById(result.itemData);
            if (file) {
                const fileid = file.id.toString().replace(/^\d+_(.*)$/g, "$1");
                const temp = ServiceFactory.getItemByName(result.itemType);
                if(typeof(temp.id) === "number") {
                    file.id = parseInt(fileid);
                }
                else {
                    file.id = fileid;
                }
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
        const instance = ServiceFactory.getItemByName(itemTypeName);
        return (instance instanceof BaseFile);
    }

}