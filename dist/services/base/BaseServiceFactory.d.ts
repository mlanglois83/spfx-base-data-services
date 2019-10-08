import { BaseDataService } from "./BaseDataService";
import { IBaseItem } from "../../interfaces";
export declare class BaseServiceFactory {
    create<T extends IBaseItem>(serviceName: string): BaseDataService<T>;
    getItemTypeByName(typeName: any): (new (item?: any) => IBaseItem);
}
