export class SPFile {
    public content?: ArrayBuffer;
    public mimeType: string; 
    public id: string;
    public title: string;


    public get serverRelativeUrl(): string {
        return this.id;
    }
    public set serverRelativeUrl(val: string) {
        this.id = val;
    }

    public get name(): string {
        return this.title;
    }
    public set name(val: string) {
        this.title = val;
    }

    constructor(fileItem?:any){
        if(fileItem) {
            this.serverRelativeUrl = (fileItem.FileRef ? fileItem.FileRef : fileItem.ServerRelativeUrl);
            this.name = (fileItem.FileLeafRef ? fileItem.FileLeafRef : fileItem.Name);
        }
    }


}