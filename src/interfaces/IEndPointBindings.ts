export interface IEndPointBindings {
    getAll: IEndPointBinding;
    get: IEndPointBinding;
    addOrUpdateItem: IEndPointBinding;
    addOrUpdateItems: IEndPointBinding;
    deleteItem: IEndPointBinding;
    getItemById: IEndPointBinding;
}
export interface IEndPointBinding {
    method: "GET" | "POST" | "PUT" | "DELETE";
    url: string;
}