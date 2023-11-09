import { BaseUserService } from "./BaseUserService";
import { Decorators } from "../../decorators";
import { Entity } from "../../models";
import { IBaseUserServiceOptions } from "../../interfaces";
const dataService = Decorators.dataService;

@dataService("Entity")
export class EntityService extends BaseUserService<Entity> {

    constructor(options?: IBaseUserServiceOptions, ...args: any[]) {
        super(Entity, options, ...args);
    }


}