import { Text } from "@microsoft/sp-core-library";
import { cloneDeep, find } from "@microsoft/sp-lodash-subset";
import { stringIsNullOrEmpty } from "@pnp/common";
import { ITerm, ITermSet, taxonomy } from "@pnp/sp-taxonomy";
import { ServiceFactory } from "../ServiceFactory";
import { UtilsService } from "../UtilsService";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { Constants } from "../../constants/index";
import { TaxonomyHidden, TaxonomyTerm } from "../../models";
import { BaseDataService } from "./BaseDataService";
import { Decorators } from "../../decorators";

const trace = Decorators.trace;
const standardTermSetCacheDuration = 10;

/**
 * Base service for sp termset operations
 */
export class BaseTermsetService<T extends TaxonomyTerm> extends BaseDataService<T> {

    protected utilsService: UtilsService;
    protected termsetnameorid: string;
    protected isGlobal: boolean;

    /**
     * Associeted termset (pnpjs)
     */
    protected get termset(): ITermSet {
        if (this.termsetnameorid.match(/[A-z0-9]{8}-([A-z0-9]{4}-){3}[A-z0-9]{12}/)) {
            if (this.isGlobal) {
                return taxonomy.getDefaultSiteCollectionTermStore().getTermSetById(this.termsetnameorid);
            }
            else {
                return taxonomy.getDefaultSiteCollectionTermStore().getSiteCollectionGroup().termSets.getById(this.termsetnameorid);
            }
        }
        else {
            if (this.isGlobal) {
                return taxonomy.getDefaultSiteCollectionTermStore().getTermSetsByName(this.termsetnameorid, 1033).getByName(this.termsetnameorid);
            }
            else {
                return taxonomy.getDefaultSiteCollectionTermStore().getSiteCollectionGroup().termSets.getByName(this.termsetnameorid);
            }
        }
    }

    protected set customSortOrder(value: string) {
        localStorage.setItem(Text.format(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName), value ? value : "");
    }
    protected get customSortOrder(): string {
        return localStorage.getItem(Text.format(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.context.pageContext.web.serverRelativeUrl, this.serviceName));
    }


    /**
     * 
     * @param type - items type
     * @param context - current sp component context 
     * @param termsetname - term set name
     */
    constructor(type: (new (item?: any) => T), termsetnameorid: string, isGlobal = true, cacheDuration: number = standardTermSetCacheDuration) {
        super(type, cacheDuration);
        this.utilsService = new UtilsService();
        this.termsetnameorid = termsetnameorid;
        this.isGlobal = isGlobal;
    }

    @trace()
    public async getWssIds(termId: string): Promise<Array<number>> {
        const taxonomyHiddenItems = await ServiceFactory.getService(TaxonomyHidden).getAll();
        return taxonomyHiddenItems.filter((taxItem) => {
            return taxItem.termId === termId;
        }).map((filteredItem) => {
            return filteredItem.id;
        });
    }

    /**
     * Retrieve all terms
     */
    @trace()
    protected async getAll_Internal(): Promise<Array<T>> {

        const batch = taxonomy.createBatch();

        let spterms: ITerm[];
        let ts: any;
        this.termset.terms.select("Name", "Description", "Id", "PathOfTerm", "CustomSortOrder", "CustomProperties", "IsDeprecated").inBatch(batch).get().then((results) => {
            spterms = results;   
        });
        this.termset.select("CustomSortOrder").inBatch(batch).get().then((result) => {
            ts = result;
        });

        await batch.execute();

        this.customSortOrder = ts.CustomSortOrder;
        const taxonomyHiddenItems = await ServiceFactory.getService(TaxonomyHidden).getAll();
        return spterms.map((term) => {
            const result = new this.itemType(term);
            result.wssids = [];
            for (const taxonomyHiddenItem of taxonomyHiddenItems) {
                if (taxonomyHiddenItem.termId == result.id) { result.wssids.push(taxonomyHiddenItem.id); }
            }

            return result;
        });
    }

    @trace()
    public async getItemById_Internal(id: string): Promise<T> {
        let result = null;
        const spterm = await this.termset.terms.getById(id).select("Name", "Description", "Id", "PathOfTerm", "CustomSortOrder", "CustomProperties", "IsDeprecated");
        if (spterm) {
            result = new this.itemType(spterm);
        }
        return result;
    }

    @trace()
    public async getItemsById_Internal(ids: Array<string>): Promise<Array<T>> {
        const results: Array<T> = [];
        const batches = [];
        const copy = cloneDeep(ids);
        while (copy.length > 0) {
            const sub = copy.splice(0, 100);
            const batch = taxonomy.createBatch();
            sub.forEach((id) => {
                this.termset.terms.getById(id).select("Name", "Description", "Id", "PathOfTerm", "CustomSortOrder", "CustomProperties", "IsDeprecated").inBatch(batch).get().then((term) => {
                    if (term) {
                        results.push(new this.itemType(term));
                    }
                    else {
                        console.log(`[${this.serviceName}] - term with id ${id} not found`);
                    }
                });
            });
            batches.push(batch);
        }
        await UtilsService.runBatchesInStacks(batches, 3);
        return results;
    }

    protected async get_Internal(query: any): Promise<Array<T>> {
        console.log("[" + this.serviceName + ".get_Internal] - " + query.toString());
        throw new Error('Not Implemented');
    }


    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        console.log("[" + this.serviceName + ".addOrUpdateItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async addOrUpdateItems_Internal(items: Array<T>/*, onItemUpdated?: (oldItem: T, newItem: T) => void*/): Promise<Array<T>> {
        console.log("[" + this.serviceName + ".addOrUpdateItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: T): Promise<T> {
        console.log("[" + this.serviceName + ".deleteItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async deleteItems_Internal(items: Array<T>): Promise<Array<T>> {
        console.log("[" + this.serviceName + ".deleteItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    @trace()
    protected async persistItemData_internal(data: any): Promise<T> {
        let result = null;
        if (data) {
            result = new this.itemType(data);
        }
        return result;
    }



    private getOrderedChildTerms(term: T, allTerms: Array<T>): Array<T> {
        //items.sort((a: T,b: T) => {return a.path.localeCompare(b.path);});
        const result = [];
        const childterms = allTerms.filter((t) => { return t.path.indexOf(term.path + ";") == 0; });
        const level = term.path.split(";").length;
        let directChilds = childterms.filter((ct) => { return ct.path.split(";").length === level + 1; });
        const terms = [];
        const orderIds = stringIsNullOrEmpty(term.customSortOrder) ? [] : term.customSortOrder.split(":");
        orderIds.forEach(id => {
            const t = find(directChilds, (spterm) => {
                return spterm.id === id;
            });
            terms.push(t);
        });
        const otherterms = directChilds.filter(spterm => !orderIds.some(o => o === spterm.id));
        otherterms.sort((a,b) => { 
            return a.title?.localeCompare(b.title); 
        });
        terms.push(...otherterms);
        directChilds = terms;
        directChilds.forEach((dc) => {
            result.push(dc);
            const dcchildren = this.getOrderedChildTerms(dc, childterms);
            if (dcchildren.length > 0) {
                result.push(...dcchildren);
            }
        });
        return result;
    }

    @trace()
    public async getAll(): Promise<Array<T>> {
        const items = await super.getAll();
        const result = [];
        let rootTerms = items.filter((item: T) => { return item.path.indexOf(";") === -1; });
        const terms = [];
        const orderIds = stringIsNullOrEmpty(this.customSortOrder) ? [] : this.customSortOrder.split(":");
        orderIds.forEach(id => {
            const term = find(rootTerms, (spterm) => {
                return spterm.id === id;
            });
            terms.push(term);
        });
        const otherterms = rootTerms.filter(spterm => !orderIds.some(o => o === spterm.id));
        otherterms.sort((a,b) => {
            return a.title?.localeCompare(b.title);
        });
        terms.push(...otherterms);
        rootTerms = terms;
        rootTerms.forEach((rt) => {
            result.push(rt);
            const rtchildren = this.getOrderedChildTerms(rt, items);
            if (rtchildren.length > 0) {
                result.push(...rtchildren);
            }
        });
        return result;
    }
}
