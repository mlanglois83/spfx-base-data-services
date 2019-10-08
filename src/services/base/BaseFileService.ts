import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ChunkedFileUploadProgressData, Folder, sp } from "@pnp/sp";
import * as mime from "mime-types";
import { UtilsService } from "../";
import { IDataService } from "../../interfaces/IDataService";
import { IBaseItem } from "../../interfaces/index";
import { SPFile } from "../../models";
import { BaseDataService } from "./BaseDataService";
import { BaseService } from "./BaseService";

/**
 * Base service for sp files operations
 */
export class BaseFileService<T extends IBaseItem> extends BaseDataService<T> implements IDataService<T>{
    protected itemType: (new (item?: any) => T);
    protected listRelativeUrl: string;




    /**
     * Associeted list (pnpjs)
     */
    protected get list() {
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
        this.itemType = type;
        this.listRelativeUrl = BaseService.Configuration.context.pageContext.web.serverRelativeUrl + listRelativeUrl;
    }
    /**
     * Retrieve all items
     */
    public async getAll_Internal(): Promise<Array<T>> {
        let files = await this.list.items.filter('FSObjType eq 0').select('FileRef', 'FileLeafRef').get();
        return await Promise.all(files.map((file) => {
            return this.createFileObject(file);
        }));
    }

    public async get_Internal(query: any): Promise<Array<T>> {

        throw new Error('Not Implemented');
    }

    public async getById_Internal(query: any): Promise<T> {

        throw new Error('Not Implemented');
    }


    private async createFileObject(file: any): Promise<T> {
        let resultFile = new this.itemType(file);
        if (resultFile instanceof SPFile) {
            resultFile.mimeType = <string>mime.lookup(resultFile.name) || 'application/octet-stream';
            //resultFile.content = await sp.web.getFileByServerRelativeUrl(resultFile.serverRelativeUrl).getBuffer();
        }
        return resultFile;
    }

    public async getFilesInFolder(folderListRelativeUrl): Promise<Array<T>> {
        let result = new Array<T>();
        const folderUrl = this.listRelativeUrl + folderListRelativeUrl;
        let folderExists = await this.folderExists(folderListRelativeUrl);
        if (folderExists) {
            let files = await sp.web.getFolderByServerRelativeUrl(folderUrl).files.get();
            result = await await Promise.all(files.map((file) => {
                return this.createFileObject(file);
            }));
        }

        return result;
    }

    public async folderExists(folderUrl): Promise<boolean> {
        let result: boolean = false;
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
            let folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            let folder: Folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
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
            await sp.web.getFileByServerRelativeUrl(item.serverRelativeUrl).delete();
            let folderUrl = UtilsService.getParentFolderUrl(item.serverRelativeUrl);
            let folder: Folder = sp.web.getFolderByServerRelativeUrl(folderUrl);
            let files = await folder.files.get();
            if (!files || files.length === 0) {
                await folder.delete();
            }
        }
    }
}
