import { BaseItem } from "../models";
import { BaseDataService } from "../services";

export interface IFactoryMapping {
    models: {[modelName: string]: new (item?: any) => BaseItem<string | number>};
    services: {[modelName: string]: new (...args: any[]) => BaseDataService<BaseItem<string | number>>};
    objects: {[typeName: string]: new () => any};
}