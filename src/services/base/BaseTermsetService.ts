import { taxonomy } from "@pnp/sp-taxonomy";
import { UtilsService } from "../";
import { Constants } from "../../constants/index";
import { TaxonomyTerm } from "../../models";
import { BaseDataService } from "./BaseDataService";
import { TaxonomyHiddenListService } from "../";
import { find } from "@microsoft/sp-lodash-subset";
import { Text } from "@microsoft/sp-core-library";
import { stringIsNullOrEmpty } from "@pnp/common";
import { BaseService } from "./BaseService";


const standardTermSetCacheDuration: number = 10;

/**
 * Base service for sp termset operations
 */
export class BaseTermsetService<T extends TaxonomyTerm> extends BaseDataService<T> {

    protected taxonomyHiddenListService: TaxonomyHiddenListService;
    protected utilsService: UtilsService;
    protected itemType: (new (item?: any) => T);
    protected termsetnameorid: string;
    protected wssIds: any = null;

    /**
     * Associeted termset (pnpjs)
     */
    protected get termset() {
        if (this.termsetnameorid.match(/[A-z0-9]{8}-([A-z0-9]{4}-){3}[A-z0-9]{12}/)) {
            return taxonomy.getDefaultSiteCollectionTermStore().getTermSetById(this.termsetnameorid);
        }
        else {
            return taxonomy.getDefaultSiteCollectionTermStore().getTermSetsByName(this.termsetnameorid, 1033).getByName(this.termsetnameorid);
        }
    }

    protected set customSortOrder(value: string) {
        localStorage.setItem(Text.format(Constants.cacheKeys.termsetCustomOrder, BaseService.Configuration.context.pageContext.web.serverRelativeUrl, this.serviceName), value ? value : "");
    }
    protected get customSortOrder(): string {
        return localStorage.getItem(Text.format(Constants.cacheKeys.termsetCustomOrder, BaseService.Configuration.context.pageContext.web.serverRelativeUrl, this.serviceName));
    }


    /**
     * 
     * @param type items type
     * @param context current sp component context 
     * @param termsetname termset name
     */
    constructor(type: (new (item?: any) => T), termsetnameorid: string, tableName: string, cacheDuration: number = standardTermSetCacheDuration) {
        super(type, tableName, cacheDuration);
        this.utilsService = new UtilsService();
        this.taxonomyHiddenListService = new TaxonomyHiddenListService();
        this.termsetnameorid = termsetnameorid;
        this.itemType = type;
    }

    public async getWssIds(termId: string): Promise<Array<number>> {
        let taxonomyHiddenItems = await this.taxonomyHiddenListService.getAll();
        return taxonomyHiddenItems.filter((taxItem) => {
            return taxItem.termId === termId;
        }).map((filteredItem) => {
            return filteredItem.id;
        });
    }

    /**
     * Retrieve all terms
     */
    protected async getAll_Internal(): Promise<Array<T>> {
        let spterms = await this.termset.terms.get();
        let ts = await this.termset.get();
        this.customSortOrder = ts.CustomSortOrder;
        let taxonomyHiddenItems = await this.taxonomyHiddenListService.getAll();
        return spterms.map((term) => {
            let result = new this.itemType(term);
            result.wssids = [];
            for (let taxonomyHiddenItem of taxonomyHiddenItems) {
                if (taxonomyHiddenItem.termId == result.id) { result.wssids.push(taxonomyHiddenItem.id); }
            }

            return result;
        });
    }

    public async getById_Internal(query: any): Promise<T> {

        throw new Error('Not Implemented');
    }

    protected async get_Internal(query: any): Promise<Array<T>> {
        throw new Error('Not Implemented');
    }


    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: T): Promise<void> {
        throw new Error("Not implemented");
    }


    private getOrderedChildTerms(term: T, allTerms: Array<T>): Array<T> {
        //items.sort((a: T,b: T) => {return a.path.localeCompare(b.path);});
        let result = [];
        let childterms = allTerms.filter((t) => { return t.path.indexOf(term.path) == 0; });
        let level = term.path.split(";").length;
        let directChilds = childterms.filter((ct) => { return ct.path.split(";").length === level + 1; });
        if (!stringIsNullOrEmpty(term.customSortOrder)) {
            let terms = new Array();
            let orderIds = term.customSortOrder.split(":");
            orderIds.forEach(id => {
                let t = find(directChilds, (spterm) => {
                    return spterm.id === id;
                });
                terms.push(t);
            });
            directChilds = terms;
        }
        directChilds.forEach((dc) => {
            result.push(dc);
            let dcchildren = this.getOrderedChildTerms(dc, childterms);
            if (dcchildren.length > 0) {
                result.push(...dcchildren);
            }
        });
        return result;
    }

    public async getAll(): Promise<Array<T>> {
        let items = await super.getAll();
        let result = [];
        let rootTerms = items.filter((item: T) => { return item.path.indexOf(";") === -1; });
        if (!stringIsNullOrEmpty(this.customSortOrder)) {
            let terms = new Array();
            let orderIds = this.customSortOrder.split(":");
            orderIds.forEach(id => {
                let term = find(rootTerms, (spterm) => {
                    return spterm.id === id;
                });
                terms.push(term);
            });
            rootTerms = terms;
        }
        rootTerms.forEach((rt) => {
            result.push(rt);
            let rtchildren = this.getOrderedChildTerms(rt, items);
            if (rtchildren.length > 0) {
                result.push(...rtchildren);
            }
        });
        return result;
    }
}
