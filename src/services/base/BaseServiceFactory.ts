import { BaseComponentContext } from "@microsoft/sp-component-base";
import { BaseDataService } from "./BaseDataService";
import { IBaseItem } from "../../interfaces";
import { SPFile } from "../../models";
export class BaseServiceFactory {
    public create<T extends IBaseItem>(serviceName: string): BaseDataService<T> {
        return null;
    }

    public getItemTypeByName(typeName): (new (item?: any) => IBaseItem) {
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
