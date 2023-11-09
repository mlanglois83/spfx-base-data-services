import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IFolder } from "@pnp/sp/folders";
import "@pnp/sp/items/list";
import "@pnp/sp/lists";
import { IList } from "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import { cloneDeep } from "lodash";
import * as mime from "mime-types";
import { TraceLevel } from "../../constants";
import { Decorators } from "../../decorators";
import { ChunkedFileUploadProgressData, IBaseSPFileServiceOptions } from "../../interfaces";
import { SPFile } from "../../models/base/SPFile";
import { UtilsService } from "../UtilsService";
import { BaseSPService } from "./BaseSPService";


const trace = Decorators.trace;
/**
 * Base service for sp files operations
 */
export class BaseFileService<T extends SPFile> extends BaseSPService<T>{
    
    protected serviceOptions: IBaseSPFileServiceOptions;
    
    protected listRelativeUrl: string;

    /**
     * Associeted list (pnpjs)
     */
    protected get list(): IList {
        return this.sp.web.getList(this.listRelativeUrl);
    }

    /**
     * 
     * @param type - items type
     * @param context - current sp component context 
     * @param listRelativeUrl - list web relative url
     */
    constructor(itemType: (new (item?: any) => T), listRelativeUrl: string, options?: IBaseSPFileServiceOptions, ...args: any[]) {
        super(itemType, options, listRelativeUrl, ...args);
        this.listRelativeUrl = this.baseRelativeUrl + listRelativeUrl;
    }
    /**
     * Retrieve all items
     */
    @trace(TraceLevel.Queries)
    public async getAll_Query(): Promise<Array<any>> {
        return this.list.items.filter('FSObjType eq 0').select('FileRef', 'FileLeafRef')();        
    }

    @trace(TraceLevel.Queries)
    public async get_Query(query: any): Promise<Array<any>> {// eslint-disable-line @typescript-eslint/no-unused-vars
        throw new Error('Not Implemented');
    }

    @trace(TraceLevel.Queries)
    public async getItemById_Query(id: string): Promise<any> {
        return this.sp.web.getFileByServerRelativePath(id)();
    }

    @trace(TraceLevel.Queries)
    public async getItemsById_Query(ids: Array<string>): Promise<Array<any>> {
        const results: Array<any> = [];
        const batches = [];
        const copy = cloneDeep(ids);
        while(copy.length > 0) {
            const sub = copy.splice(0,100);
            const [batchedSP, execute]= this.sp.batched();
            sub.forEach((id) => {
                batchedSP.web.getFileByServerRelativePath(id)().then(async (item)=> {
                    if(item) {                        
                        results.push(item);
                    }
                    else {                        
                        console.log(`[${this.serviceName}] - file with url ${id} not found`);
                    }
                });
            });
            batches.push(execute);
        }    
        await UtilsService.runBatchesInStacks(batches, 3);    
        return results;
    }

    protected populateItem(file: any): T {
        const resultFile = super.populateItem(file);
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
            const files = await this.sp.web.getFolderByServerRelativePath(folderUrl).files();
            result = files.map(file => this.populateItem(file));
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
            await this.sp.web.getFolderByServerRelativePath(folderUrl)();
            result = true;
        } catch (error) {
            // no folder, returns false
        }
        return result;
    }

    @trace(TraceLevel.Internal)
    public async addOrUpdateItem_Internal(item: T): Promise<T> {
        const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
        const folder: IFolder = this.sp.web.getFolderByServerRelativePath(folderUrl);
        const exists = await this.folderExists(folderUrl);
        if (!exists) {
            await this.sp.web.folders.addUsingPath(folderUrl);
        }
        if (item.content.byteLength <= 10485760) {
            // small upload
            try {
                await folder.files.addUsingPath(item.title, item.content, {Overwrite: true});
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
                onItemUpdated(item as T, item as T);
            }     
        });
        return items;
    }

    @trace(TraceLevel.Internal)
    public async deleteItem_Internal(item: T): Promise<T> {        
        if(item.id) {
            await this.sp.web.getFileByServerRelativePath(item.serverRelativeUrl).delete();
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            const folder: IFolder = this.sp.web.getFolderByServerRelativePath(folderUrl);
            const files = await folder.files();
            if (!files || files.length === 0) {
                await folder.delete();
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
        const [batchedSP, execute] = this.sp.batched();   
        const folders = [];
        items.filter(i => i.id).forEach(item => {
            batchedSP.web.getFileByServerRelativePath(item.serverRelativeUrl).delete().then(() => {
                item.deleted = true;
            }).catch((error) => {
                item.error = error;
            });
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            if(folders.indexOf(folderUrl) === -1) {
                folders.push(folderUrl);
            }
        });  
        await execute();
        const [folderbatchedSP, folderExecute] = this.sp.batched();
        folders.forEach(f => {
            folderbatchedSP.web.getFolderByServerRelativePath(f).files().then(async (files) => {
                if (!files || files.length === 0) {
                    await this.sp.web.getFolderByServerRelativePath(f).delete();
                } 
            });
        });   
        await folderExecute();
        return items;
    }

    @trace(TraceLevel.Internal)
    public async recycleItem_Internal(item: T): Promise<T> {        
        if(item.id) {
            await this.sp.web.getFileByServerRelativePath(item.serverRelativeUrl).recycle();
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            const folder: IFolder = this.sp.web.getFolderByServerRelativePath(folderUrl);
            const files = await folder.files();
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
    public async recycleItems_Internal(items: Array<T>): Promise<Array<T>> { 
        items.filter(i => !i.id).forEach(i => i.deleted = true);   
        const [batchedSP, execute] = this.sp.batched();   
        const folders = [];
        items.filter(i => i.id).forEach(item => {
            batchedSP.web.getFileByServerRelativePath(item.serverRelativeUrl).recycle().then(() => {
                item.deleted = true;
            }).catch((error) => {
                item.error = error;
            });
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            if(folders.indexOf(folderUrl) === -1) {
                folders.push(folderUrl);
            }
        });  
        await execute();
        const [folderbatchedSP, folderExecute] = this.sp.batched();
        folders.forEach(f => {
            folderbatchedSP.web.getFolderByServerRelativePath(f).files().then(async (files) => {
                if (!files || files.length === 0) {
                    await this.sp.web.getFolderByServerRelativePath(f).recycle();
                } 
            });
        });   
        await folderExecute();
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
