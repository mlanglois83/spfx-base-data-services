import { UserService } from "..";
import { Decorators } from "../../decorators";
const standardUserCacheDuration = 10;
const dataService = Decorators.dataService;



@dataService("Entity")
export class EntityService extends UserService {

    constructor(cacheDuration: number = standardUserCacheDuration) {
        super(cacheDuration, true);
    }


}