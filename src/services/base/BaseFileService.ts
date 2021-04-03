import { ChunkedFileUploadProgressData, Folder, sp, List } from "@pnp/sp";
import * as mime from "mime-types";
import { UtilsService } from "../UtilsService";
import { BaseDataService } from "./BaseDataService";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { SPFile } from "../../models/base/SPFile";
import { Decorators } from "../../decorators";
import { TraceLevel } from "../../constants";

const trace = Decorators.trace;
/**
 * Base service for sp files operations
 */
export class BaseFileService<T extends SPFile> extends BaseDataService<T>{
    protected listRelativeUrl: string;

    /**
     * Associeted list (pnpjs)
     */
    protected get list(): List {
        return sp.web.getList(this.listRelativeUrl);
    }

    /**
     * 
     * @param type - items type
     * @param context - current sp component context 
     * @param listRelativeUrl - list web relative url
     */
    constructor(type: (new (item?: any) => T), listRelativeUrl: string) {
        super(type);
        this.listRelativeUrl = ServicesConfiguration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
    }
    /**
     * Retrieve all items
     */
    @trace(TraceLevel.Queries)
    public async getAll_Query(): Promise<Array<any>> {
        return this.list.items.filter('FSObjType eq 0').select('FileRef', 'FileLeafRef').get();        
    }

    @trace(TraceLevel.Queries)
    public async get_Query(query: any): Promise<Array<any>> {// eslint-disable-line @typescript-eslint/no-unused-vars
        throw new Error('Not Implemented');
    }

    @trace(TraceLevel.Queries)
    public async getItemById_Query(id: string): Promise<any> {
        return sp.web.getFileByServerRelativeUrl(id).select('FileRef', 'FileLeafRef').get();
    }

    @trace(TraceLevel.Queries)
    public async getItemsById_Query(ids: Array<string>): Promise<Array<any>> {
        const results: Array<any> = [];
        const batches = [];
        const copy = cloneDeep(ids);
        while(copy.length > 0) {
            const sub = copy.splice(0,100);
            const batch = sp.createBatch();
            sub.forEach((id) => {
                sp.web.getFileByServerRelativeUrl(id).select('FileRef', 'FileLeafRef').get().then(async (item)=> {
                    if(item) {                        
                        results.push(item);
                    }
                    else {                        
                        console.log(`[${this.serviceName}] - file with url ${id} not found`);
                    }
                });
            });
            batches.push(batch);
        }    
        await UtilsService.runBatchesInStacks(batches, 3);    
        return results;
    }

    protected async populateItem(file: any): Promise<T> {
        const resultFile = new this.itemType(file);
        resultFile.mimeType = (mime.lookup(resultFile.title) as string) || 'application/octet-stream';
        return resultFile;
    }
    protected async convertItem(item: T): Promise<any> {// eslint-disable-line @typescript-eslint/no-unused-vars
        throw Error("Not implemented");
    }


    @trace(TraceLevel.Service)
    public async getFilesInFolder(folderListRelativeUrl): Promise<Array<T>> {
        let result = new Array<T>();
        const folderUrl = this.listRelativeUrl + folderListRelativeUrl;
        const folderExists = await this.folderExists(folderListRelativeUrl);
        if (folderExists) {
            const files = await sp.web.getFolderByServerRelativeUrl(folderUrl).files.get();
            result = await await Promise.all(files.map((file) => {
                return this.populateItem(file);
            }));
        }

        return result;
    }

    @trace(TraceLevel.Service)
    public async folderExists(folderUrl): Promise<boolean> {
        let result = false;
        if (folderUrl.indexOf(this.listRelativeUrl) === -1) {
            folderUrl = this.listRelativeUrl + folderUrl;
        }
        try {
            await sp.web.getFolderByServerRelativeUrl(folderUrl).get();
            result = true;
        } catch (error) {
            // no folder, returns false
        }
        return result;
    }

    @trace(TraceLevel.Internal)
    public async addOrUpdateItem_Internal(item: T): Promise<T> {
        const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
        const folder: Folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
        const exists = await this.folderExists(folderUrl);
        if (!exists) {
            await sp.web.folders.add(folderUrl);
        }
        if (item.content.byteLength <= 10485760) {
            // small upload
            try {
                await folder.files.add(item.title, item.content, true);
            }
            catch(error) {
                item.error = error;
            }
        } else {
            // large upload
            try {
                await folder.files.addChunked(item.title, UtilsService.arrayBufferToBlob(item.content, item.mimeType), (data: ChunkedFileUploadProgressData) => {
                    console.log("block:" + data.blockNumber + "/" + data.totalBlocks);
                }, true);
            }
            catch(error) {
                item.error = error;
            }            
        }
        return item;
    }

    @trace(TraceLevel.Internal)
    public async addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        const result = [];
        const operations = items.map((item) => {
            return this.addOrUpdateItem_Internal(item);
        });
        operations.reduce((chain, operation) => {                  
            return chain.then(() => {return operation;});                  
        }, Promise.resolve()).then((item) => {
            result.push(item);       
            if(onItemUpdated) {
                onItemUpdated(item, item);
            }     
        });
        return items;
    }

    @trace(TraceLevel.Internal)
    public async deleteItem_Internal(item: T): Promise<T> {        
        if(item.id) {
            await sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl).recycle();
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            const folder: Folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
            const files = await folder.files.get();
            if (!files || files.length === 0) {
                await folder.recycle();
            }
            item.deleted = true;
        }
        else {
            item.deleted = true;
        }        
        return item;
    }

    @trace(TraceLevel.Internal)
    public async deleteItems_Internal(items: Array<T>): Promise<Array<T>> { 
        items.filter(i => !i.id).forEach(i => i.deleted = true);   
        const batch = sp.createBatch();   
        const folders = [];
        items.filter(i => i.id).forEach(item => {
            sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl).inBatch(batch).recycle().then(() => {
                item.deleted = true;
            }).catch((error) => {
                item.error = error;
            });
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            if(folders.indexOf(folderUrl) === -1) {
                folders.push(folderUrl);
            }
        });  
        await batch.execute();
        const folderbatch = sp.createBatch();
        folders.forEach(f => {
            sp.web.getFolderByServerRelativeUrl(f).files.inBatch(folderbatch).get().then(async (files) => {
                if (!files || files.length === 0) {
                    await sp.web.getFolderByServerRelativeUrl(f).recycle();
                } 
            });
        });   
        return items;
    }
    
    @trace(TraceLevel.Service)
    public async changeFolderInDb(oldFolderListRelativeUrl: string, newFolderListRelativeUrl: string): Promise<void> {
        const oldFolderRelativeUrl = this.listRelativeUrl + oldFolderListRelativeUrl;
        const newFolderRelativeUrl = this.listRelativeUrl + newFolderListRelativeUrl;

        const allFiles = await this.dbService.getAll();
        const files = allFiles.filter(f => {
            return UtilsService.getParentFolderUrl(f.id.toString()).toLowerCase() === oldFolderRelativeUrl.toLowerCase();
        });
        const newFiles = cloneDeep(files);
        await Promise.all(files.map((f) => {
            return this.dbService.deleteItem(f);
        }));
        newFiles.forEach((file) => {
            file.id = newFolderRelativeUrl + "/" + file.title;            
        });
        await this.dbService.addOrUpdateItems(newFiles);
    }
}
