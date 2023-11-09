
export interface IBaseDataServiceOptions {
    cacheDuration?: number;
}

export interface IBaseSPServiceOptions extends IBaseDataServiceOptions {
    baseUrl?: string;
}

// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface IBaseSPContainerServiceOptions extends IBaseSPServiceOptions {
}
// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface IBaseSPFileServiceOptions extends IBaseSPContainerServiceOptions {
}

export interface IBaseTermsetServiceOptions extends IBaseSPContainerServiceOptions {
    isGlobal?: boolean;
}

export interface IBaseListItemServiceOptions extends IBaseSPContainerServiceOptions {
    useOData?: boolean;
    multiSite?: boolean;
}

// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface IBaseRestServiceOptions extends IBaseDataServiceOptions {
}

export interface IBaseUserServiceOptions extends IBaseSPServiceOptions {
    includeGroups?: boolean;
}