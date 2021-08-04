import { Text } from "@microsoft/sp-core-library";
import { cloneDeep, find } from "@microsoft/sp-lodash-subset";
import { stringIsNullOrEmpty } from "@pnp/common";
import { ITermSet, taxonomy } from "@pnp/sp-taxonomy";
import { ServiceFactory } from "../ServiceFactory";
import { UtilsService } from "../UtilsService";
import { ServicesConfiguration } from "../../configuration/ServicesConfiguration";
import { Constants, TraceLevel } from "../../constants/index";
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

    @trace(TraceLevel.ServiceUtilities)
    protected async init_internal(): Promise<void> {
        await super.init_internal();
        const [ts, taxonomyHiddenItems] = await Promise.all([this.termset.select("CustomSortOrder").get(), ServiceFactory.getService(TaxonomyHidden).getAll()]);
        this.customSortOrder = ts.CustomSortOrder;
        this.updateInitValues(TaxonomyHidden["name"], ...taxonomyHiddenItems);
    }

    protected populateItem(data: any): T {
        const result = new this.itemType(data);
        const taxonomyHiddenItems = this.getServiceInitValues(TaxonomyHidden);
        result.wssids = [];
        for (const taxonomyHiddenItem of taxonomyHiddenItems) {
            if (taxonomyHiddenItem.termId == result.id) { result.wssids.push(taxonomyHiddenItem.id); }
        }
        return result;
    }

    protected async convertItem(item: T): Promise<any> {// eslint-disable-line @typescript-eslint/no-unused-vars
        throw Error("Not implemented");
    }

    @trace(TraceLevel.Service)
    public async getWssIds(termId: string): Promise<Array<number>> {
        if (!this.initialized) {
            await this.Init();
        }
        const taxonomyHiddenItems = this.getServiceInitValues(TaxonomyHidden);
        return taxonomyHiddenItems.filter((taxItem) => {
            return taxItem.termId === termId;
        }).map((filteredItem) => {
            return filteredItem.id;
        });
    }
    @trace(TraceLevel.Queries)
    protected async getAll_Query(): Promise<Array<any>> {
        return this.termset.terms.select("Name", "Description", "Id", "PathOfTerm", "CustomSortOrder", "CustomProperties", "IsDeprecated").get();
    }
    

    @trace(TraceLevel.Queries)
    public async getItemById_Query(id: string): Promise<any> {
        return  this.termset.terms.getById(id).select("Name", "Description", "Id", "PathOfTerm", "CustomSortOrder", "CustomProperties", "IsDeprecated");
    }

    @trace(TraceLevel.Queries)
    public async getItemsById_Query(ids: Array<string>): Promise<Array<any>> {
        const results: Array<any> = [];
        const batches = [];
        const copy = cloneDeep(ids);
        while (copy.length > 0) {
            const sub = copy.splice(0, 100);
            const batch = taxonomy.createBatch();
            sub.forEach((id) => {
                this.termset.terms.getById(id).select("Name", "Description", "Id", "PathOfTerm", "CustomSortOrder", "CustomProperties", "IsDeprecated").inBatch(batch).get().then((term) => {
                    if (term) {
                        results.push(term);
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

    protected async get_Query(query: any): Promise<Array<any>> { // eslint-disable-line @typescript-eslint/no-unused-vars
      throw new Error("Not Implemented");
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

    @trace(TraceLevel.Service)
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
