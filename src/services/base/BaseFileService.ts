import { ChunkedFileUploadProgressData, Folder, sp, List } from "@pnp/sp";
import * as mime from "mime-types";
import { UtilsService } from "../";
import { IBaseItem } from "../../interfaces/index";
import { SPFile } from "../../models";
import { BaseDataService } from "./BaseDataService";
import { ServicesConfiguration } from "../..";
import { cloneDeep } from "@microsoft/sp-lodash-subset";

/**
 * Base service for sp files operations
 */
export class BaseFileService<T extends IBaseItem> extends BaseDataService<T>{
    protected listRelativeUrl: string;

    /**
     * Associeted list (pnpjs)
     */
    protected get list(): List {
        return sp.web.getList(this.listRelativeUrl);
    }

    /**
     * 
     * @param type items type
     * @param context current sp component context 
     * @param listRelativeUrl list web relative url
     */
    constructor(type: (new (item?: any) => T), listRelativeUrl: string, tableName: string) {
        super(type, tableName);
        this.listRelativeUrl = ServicesConfiguration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
    }
    /**
     * Retrieve all items
     */
    public async getAll_Internal(): Promise<Array<T>> {
        const files = await this.list.items.filter('FSObjType eq 0').select('FileRef', 'FileLeafRef').get();
        return await Promise.all(files.map((file) => {
            return this.createFileObject(file);
        }));
    }

    public async get_Internal(query: any): Promise<Array<T>> {
        console.log("[" + this.serviceName + ".get_Internal] - " + query.toString());
        throw new Error('Not Implemented');
    }

    public async getItemById_Internal(id: string): Promise<T> {
        let result = null;
        const file = await sp.web.getFileByServerRelativeUrl(id).select('FileRef', 'FileLeafRef').get();
        if(file) {
            result = this.createFileObject(file);
        }
        return result;
    }
    public async getItemsById_Internal(ids: Array<string>): Promise<Array<T>> {
        const results: Array<T> = [];
        const batch = sp.createBatch();
        ids.forEach((id) => {
            sp.web.getFileByServerRelativeUrl(id).select('FileRef', 'FileLeafRef').get().then(async (item)=> {
                const fo = await this.createFileObject(item);
                results.push(fo);
            });
        });
        await batch.execute();
        return results;
    }

    private async createFileObject(file: any): Promise<T> {
        const resultFile = new this.itemType(file);
        if (resultFile instanceof SPFile) {
            resultFile.mimeType = (mime.lookup(resultFile.name) as string) || 'application/octet-stream';
            //resultFile.content = await sp.web.getFileByServerRelativeUrl(resultFile.serverRelativeUrl).getBuffer();
        }
        return resultFile;
    }

    public async getFilesInFolder(folderListRelativeUrl): Promise<Array<T>> {
        let result = new Array<T>();
        const folderUrl = this.listRelativeUrl + folderListRelativeUrl;
        const folderExists = await this.folderExists(folderListRelativeUrl);
        if (folderExists) {
            const files = await sp.web.getFolderByServerRelativeUrl(folderUrl).files.get();
            result = await await Promise.all(files.map((file) => {
                return this.createFileObject(file);
            }));
        }

        return result;
    }

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

    public async addOrUpdateItem_Internal(item: T): Promise<T> {
        if (item instanceof SPFile && item.content) {
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            const folder: Folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
            const exists = await this.folderExists(folderUrl);
            if (!exists) {
                await sp.web.folders.add(folderUrl);
            }
            if (item.content.byteLength <= 10485760) {
                // small upload
                await folder.files.add(item.name, item.content, true);
            } else {
                // large upload
                await folder.files.addChunked(item.name, UtilsService.arrayBufferToBlob(item.content, item.mimeType), (data: ChunkedFileUploadProgressData) => {
                    console.log("block:" + data.blockNumber + "/" + data.totalBlocks);
                }, true);
            }
        }
        return item;
    }

    public async deleteItem_Internal(item: T): Promise<void> {
        if (item instanceof SPFile) {
            await sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl).recycle();
            const folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            const folder: Folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
            const files = await folder.files.get();
            if (!files || files.length === 0) {
                await folder.recycle();
            }
        }
    }
    
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
