export declare class SPFile {
    content?: ArrayBuffer;
    mimeType: string;
    id: string;
    title: string;
    get serverRelativeUrl(): string;
    set serverRelativeUrl(val: string);
    get name(): string;
    set name(val: string);
    constructor(fileItem?: any);
}
