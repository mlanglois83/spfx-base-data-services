import { Decorators } from "../../decorators";
import { User } from "..";
const dataModel = Decorators.dataModel;


@dataModel()
export class Entity extends User {


    constructor(userObj?: any) {
        super(userObj);

    }

}