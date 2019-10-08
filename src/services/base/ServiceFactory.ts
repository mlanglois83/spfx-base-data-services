import { BaseComponentContext } from "@microsoft/sp-component-base";
import { BaseDataService } from "./BaseDataService";
import { IBaseItem } from "../../interfaces";
import { SPFile } from "../../models";
export abstract class ServiceFactory {

    public static create<T extends IBaseItem>(context: BaseComponentContext, serviceName: string): BaseDataService<T> {
        return null;
    }

    public static getItemTypeByName(typeName): (new (item?: any) => IBaseItem) {
        let result = null;
        switch (typeName) {
            case SPFile["name"]:
                result = SPFile;
                break;
            default:
                break;
        }
        return result;
    }

}