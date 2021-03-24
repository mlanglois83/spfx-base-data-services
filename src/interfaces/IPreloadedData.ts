export interface IPreloadedData{
    [modelName: string]: {
        [itemId: string]: {
            [propertyName: string]: any;
        };
    };
}