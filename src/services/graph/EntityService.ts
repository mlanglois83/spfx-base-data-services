import { BaseUserService } from "./BaseUserService";
import { Decorators } from "../../decorators";
import { Entity } from "../../models";
const standardUserCacheDuration = 10;
const dataService = Decorators.dataService;

@dataService("Entity")
export class EntityService extends BaseUserService<Entity> {

    constructor(cacheDuration: number = standardUserCacheDuration, baseUrl?: string, ...args: any[]) {
        super(Entity, cacheDuration, true, baseUrl, ...args);
    }


}