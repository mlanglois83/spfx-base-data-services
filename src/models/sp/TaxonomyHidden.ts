import { SPItem } from "../";


export class TaxonomyHidden extends SPItem {
    public id: number;
    public termId: string;

    constructor(item: any) {
        super(item);

        if (item != undefined) {
            this.termId = item.IdForTerm;
        }
    }

    public convert(): Promise<any> {
        throw new Error("Not implemented");
    }
}