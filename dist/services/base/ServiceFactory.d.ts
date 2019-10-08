import { BaseComponentContext } from "@microsoft/sp-component-base";
import { BaseDataService } from "./BaseDataService";
import { IBaseItem } from "../../interfaces";
export declare abstract class ServiceFactory {
    static create<T extends IBaseItem>(context: BaseComponentContext, serviceName: string): BaseDataService<T>;
    static getItemTypeByName(typeName: any): (new (item?: any) => IBaseItem);
}
