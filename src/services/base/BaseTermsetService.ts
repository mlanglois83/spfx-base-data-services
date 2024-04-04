import { dateAdd, PnPClientStorage, stringIsNullOrEmpty } from "@pnp/core";
import "@pnp/sp/sites";
import { IOrderedTermInfo, ITermSet } from "@pnp/sp/taxonomy";
import "@pnp/sp/webs";
import { find, findIndex } from "lodash";
import { ServicesConfiguration } from "../../configuration";
import { Constants, TraceLevel } from "../../constants/index";
import { Decorators } from "../../decorators";
import { IBaseTermsetServiceOptions } from "../../interfaces";
import { BaseItem, TaxonomyHidden, TaxonomyTerm } from "../../models";
import "../../pnpExtensions/TermsetExt";
import { ServiceFactory } from "../ServiceFactory";
import { UtilsService } from "../UtilsService";
import { BaseSPService } from "./BaseSPService";
const trace = Decorators.trace;
const standardTermSetCacheDuration = 10;

/**
 * Base service for sp termset operations
 */
export class BaseTermsetService<
    T extends TaxonomyTerm
    > extends BaseSPService<T> {
    protected utilsService: UtilsService;
    protected static siteCollectionTermsetId: string;
    protected serviceOptions: IBaseTermsetServiceOptions;

    protected termsetnameorid: string;
    
    protected set customSortOrder(value: string) {
        localStorage.setItem(UtilsService.formatText(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.configuration.serviceKey, this.cacheKeyUrl, this.serviceName), value ? value : "");
    }
    protected get customSortOrder(): string {
        return localStorage.getItem(UtilsService.formatText(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.configuration.serviceKey, this.cacheKeyUrl, this.serviceName));
    }

    protected set tsId(value: string) {
        localStorage.setItem(UtilsService.formatText(Constants.cacheKeys.termsetId, ServicesConfiguration.configuration.serviceKey, this.cacheKeyUrl, this.serviceName), value ? value : "");
    }
    protected get tsId(): string {
        return localStorage.getItem(UtilsService.formatText(Constants.cacheKeys.termsetId, ServicesConfiguration.configuration.serviceKey, this.cacheKeyUrl, this.serviceName));
    }

    protected set siteCollectionGroupId(value: string) {
        localStorage.setItem(UtilsService.formatText(Constants.cacheKeys.termsetSiteCollectionGroupId, ServicesConfiguration.configuration.serviceKey, this.cacheKeyUrl, this.serviceName), value ? value : "");
    }
    protected get siteCollectionGroupId(): string {
        return localStorage.getItem(UtilsService.formatText(Constants.cacheKeys.termsetSiteCollectionGroupId, ServicesConfiguration.configuration.serviceKey, this.cacheKeyUrl, this.serviceName));
    }

    protected static setTermStoreDefaultLanguageTag(value: string, cacheUrl: string) {
        localStorage.setItem(UtilsService.formatText(Constants.cacheKeys.termStoreDefaultLanguageTag, ServicesConfiguration.configuration.serviceKey, cacheUrl), value ? value : "");
    }
    protected static getTermStoreDefaultLanguageTag(cacheUrl: string): string {
        return localStorage.getItem(UtilsService.formatText(Constants.cacheKeys.termStoreDefaultLanguageTag, ServicesConfiguration.configuration.serviceKey, cacheUrl));
    }

     /**************************************** Taxo hidden list storage ****************************************************/
    protected initValues: { [modelName: string]: BaseItem<string | number>[] } = {};
    
    protected getServiceInitValues<Tvalue extends BaseItem<string | number>>(model: new (data?: any) => Tvalue): Tvalue[] {
        return this.getServiceInitValuesByName<Tvalue>(model["name"]);
    }

    protected getServiceInitValuesByName<Tvalue extends BaseItem<string | number>>(modelName: string): Tvalue[] {
        return this.initValues[modelName] as Tvalue[];
    }

    protected updateInitValues(modelName: string, ...items: BaseItem<string | number>[]): void {
        this.initValues[modelName] = this.initValues[modelName] || [];
        items.forEach(i => {
            const idx = findIndex(this.initValues[modelName], iv => iv.id === i.id);
            if (idx !== -1) {
                this.initValues[modelName][idx] = i;
            }
            else {
                this.initValues[modelName].push(i);
            }
        });
    }

    /**********************************************************************************************************************/

    /**
     * Get site collection group
     */
    protected async getSiteCollectionGroupId(): Promise<string> {
        return this.callAsyncWithPromiseManagement(async () => {
            if (stringIsNullOrEmpty(this.siteCollectionGroupId)) {
                const [ts, properties] = await Promise.all([
                    this.sp.termStore(),
                    this.sp.site.rootWeb.allProperties(),
                ]);
                this.siteCollectionGroupId =
                    properties["SiteCollectionGroupId" + ts.id] ||
                    properties[
                    "SiteCollectionGroupId" + ts.id.replace(/-/g, "_x002d_")
                    ];                        
            }
            return this.siteCollectionGroupId;            
        }, "SiteColGroupId");        
    }

    protected static async initTermStoreDefaultLanguageTag(cacheUrl: string): Promise<string> {
        return UtilsService.callAsyncWithPromiseManagement("BaseTermSetService-TermStore", async () => {
            if (stringIsNullOrEmpty(BaseTermsetService.getTermStoreDefaultLanguageTag(cacheUrl))) {                
                const ts = await ServicesConfiguration.sp.termStore();
                BaseTermsetService.setTermStoreDefaultLanguageTag(ts.defaultLanguageTag, cacheUrl);
            }
            return BaseTermsetService.getTermStoreDefaultLanguageTag(cacheUrl);
        });        
    }

    /**
     * Associated termset (pnpjs)
     */
    @trace(TraceLevel.ServiceUtilities)
    protected async GetTermset(): Promise<ITermSet> {
        return this.callAsyncWithPromiseManagement(async () => {
            if (
                stringIsNullOrEmpty(this.tsId) &&
                this.termsetnameorid.match(/[A-z0-9]{8}-([A-z0-9]{4}-){3}[A-z0-9]{12}/)
            ) {
                this.tsId = this.termsetnameorid;
            }
            if (stringIsNullOrEmpty(this.tsId)) {
                if (this.serviceOptions.isGlobal) {
                    const [termsets, tsLngTag] = await Promise.all([
                        this.sp.termStore.sets(),
                        BaseTermsetService.initTermStoreDefaultLanguageTag(this.cacheKeyUrl),
                    ]);
                    const ts = find(termsets, (t) =>
                        t.localizedNames.some(
                            (ln) =>
                                ln.languageTag === tsLngTag &&
                                ln.name === this.termsetnameorid
                        )
                    );
                    if (ts) {
                        this.tsId = ts.id;
                        this.customSortOrder = ts.customSortOrder?.join(":");
                        return this.sp.termStore.sets.getById(this.tsId);
                    } else {
                        throw new Error("Termset not found: " + this.termsetnameorid);
                    }
                } else {
                    const groupId =
                        await this.getSiteCollectionGroupId();
                    const [termsets, tsLngTag] = await Promise.all([
                        this.sp.termStore.groups.getById(groupId).sets(),
                        BaseTermsetService.initTermStoreDefaultLanguageTag(this.cacheKeyUrl),
                    ]);
                    const ts = find(termsets, (t) =>
                        t.localizedNames.some(
                            (ln) =>
                                ln.languageTag === tsLngTag &&
                                ln.name === this.termsetnameorid
                        )
                    );
                    if (ts) {
                        this.tsId = ts.id;
                        this.customSortOrder = ts.customSortOrder?.join(":");
                        return this.sp.termStore.sets.getById(this.tsId);
                    } else {
                        throw new Error(
                            "Termset not found in site collection group: " +
                            this.termsetnameorid
                        );
                    }
                }
            } else {
                return this.sp.termStore.sets.getById(this.tsId);
            }
        }, "getTermSet");            
    }

    /**
     *
     * @param type - items type
     * @param context - current sp component context
     * @param termsetname - term set name
     */
    constructor(
        itemType: (new (item?: any) => T), 
        termsetIdentifier: string, 
        options?: IBaseTermsetServiceOptions,
        ...args: any[]
    ) {
        super(itemType, options, termsetIdentifier, ...args);
        this.serviceOptions = this.serviceOptions || {};
        this.serviceOptions.isGlobal = this.serviceOptions.isGlobal === undefined ? true : this.serviceOptions.isGlobal;
        this.serviceOptions.cacheDuration = this.serviceOptions.cacheDuration === undefined ? standardTermSetCacheDuration : this.serviceOptions.cacheDuration;
        this.termsetnameorid = termsetIdentifier
        this.utilsService = new UtilsService();
        
    }

    @trace(TraceLevel.ServiceUtilities)
    protected async init_internal(): Promise<void> {
        await super.init_internal();
        const [taxonomyHiddenItems] = await Promise.all([
            ServiceFactory.getService(TaxonomyHidden, {baseUrl: this.baseUrl}).getAll(),
            BaseTermsetService.initTermStoreDefaultLanguageTag(this.cacheKeyUrl),
        ]);
        this.updateInitValues(TaxonomyHidden["name"], ...taxonomyHiddenItems);
    }

    protected populateTerms(
        terms: IOrderedTermInfo[],
        basePath?: string
    ): Array<T> {
        const result = new Array<T>();
        for (const term of terms) {
            const item = this.populateTerm(term, basePath);
            result.push(item);
            if (term.childrenCount > 0) {
                result.push(...this.populateTerms(term.children as IOrderedTermInfo[], item.path));
            }
        }
        return result;
    }

    protected populateTerm(term: IOrderedTermInfo, basePath: string): T {
        const result: T = new this.itemType();
        // common properties
        result.id = term.id;
        result.isDeprecated = term.isDeprecated;
        // custom sort order
        if(term.customSortOrder && term.customSortOrder.length > 0) {
            const currentSortOrder = find(term.customSortOrder, cso => cso.setId === this.tsId);
            if(currentSortOrder) {
                result.customSortOrder = currentSortOrder.order.join(":");
            }
        }
        // translated
        result.title = this.getTermTitle(term);
        result.description = this.getTermDescription(term);
        // path
        result.path = stringIsNullOrEmpty(basePath)
            ? term.defaultLabel
            : basePath + ";" + term.defaultLabel;
        // properties
        result.customProperties = this.getTermProperties(term);
        // wssids
        const taxonomyHiddenItems = this.getServiceInitValues(TaxonomyHidden);
        result.wssids = [];
        for (const taxonomyHiddenItem of taxonomyHiddenItems) {
            if (taxonomyHiddenItem.termId == result.id) {
                result.wssids.push(taxonomyHiddenItem.id);
            }
        }
        return result;
    }

    protected getTermTitle(term: IOrderedTermInfo): string {
        return this.getTranslatedLabel(
            term.labels
                .filter((l) => l.isDefault)
                .map((l) => {
                    return { label: l.name, languageTag: l.languageTag };
                })
        );
    }

    protected getTermDescription(term: IOrderedTermInfo): string {
        return this.getTranslatedLabel(
            term.descriptions.map((l) => {
                return { label: l.description, languageTag: l.languageTag };
            })
        );
    }

    protected getTermPath(
        term: IOrderedTermInfo,
        allTerms: IOrderedTermInfo[]
    ): string {
        const parts = [this.getTermTitle(term)];
        while (term.parent) {
            term = find(allTerms, (t) => t.id === term.parent.id);
            parts.push(this.getTermTitle(term));
        }
        return parts.reverse().join(";");
    }
    protected getTermProperties(term: IOrderedTermInfo): {
        [key: string]: string;
    } {
        const result: { [key: string]: string } = {};
        if (term.properties) {
            term.properties.forEach((p) => {
                result[p.key] = p.value;
            });
        }
        return result;
    }

    protected getTranslatedLabel(
        labelCollection: { label: string; languageTag: string }[]
    ): string {
        // no context, get the current context
        // current ui language
        const current =
            ServicesConfiguration.context.pageContext.cultureInfo
                .currentUICultureName;
        const currentLabel = find(
            labelCollection,
            (label) => label.languageTag === current
        );
        if (currentLabel) {
            return currentLabel.label;
        } else {
            // web language
            const web = ServicesConfiguration.context.pageContext.web.languageName;
            const webLabel = find(
                labelCollection,
                (label) => label.languageTag === web
            );
            if (webLabel) {
                return webLabel.label;
            } else {
                // default termstore language
                const taxonomy = BaseTermsetService.getTermStoreDefaultLanguageTag(this.cacheKeyUrl);
                const taxonomyLabel = find(
                    labelCollection,
                    (label) => label.languageTag === taxonomy
                );
                if (taxonomyLabel) {
                    return taxonomyLabel.label;
                }
            }
        }
        return undefined;
    }

    @trace(TraceLevel.Service)
    public async getWssIds(termId: string): Promise<Array<number>> {
        return this.callAsyncWithPromiseManagement(async () => {
            await this.Init();
            const taxonomyHiddenItems = this.getServiceInitValues(TaxonomyHidden);
            return taxonomyHiddenItems
                .filter((taxItem) => {
                    return taxItem.termId === termId;
                })
                .map((filteredItem) => {
                    return filteredItem.id;
                });
        }, "wssIds");
    }
    @trace(TraceLevel.Queries)
    protected async getAll_Query(): Promise<Array<IOrderedTermInfo>> {
        const termset = await this.GetTermset();
        const store = new PnPClientStorage();
        return store.session.getOrPut(
            this.serviceName + "-alltermsordered",
            () => {
                return termset.getAllChildrenAsOrderedTreeFull();
            },
            dateAdd(new Date(), "minute", this.serviceOptions.cacheDuration || -1)
        );
    }

    @trace(TraceLevel.Internal)
    protected async getAll_Internal(): Promise<Array<T>> {
        let results: Array<T> = [];
        await this.Init();
        const items = await this.getAll_Query();
        if (items && items.length > 0) {
            results = this.populateTerms(items);
        }
        return results;
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
        otherterms.sort((a, b) => {
            return a.title?.localeCompare(b.title);
        });
        terms.push(...otherterms);
        rootTerms = terms;
        rootTerms.filter(Boolean).forEach((rt) => {
            result.push(rt);
            const rtchildren = this.getOrderedChildTerms(rt, items);
            if (rtchildren.length > 0) {
                result.push(...rtchildren);
            }
        });
        return result;
    }

    private getOrderedChildTerms(term: T, allTerms: Array<T>): Array<T> {
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
        otherterms.sort((a, b) => {
            return a.title?.localeCompare(b.title);
        });
        terms.push(...otherterms);
        directChilds = terms;
        directChilds.filter(Boolean).forEach((dc) => {
            result.push(dc);
            const dcchildren = this.getOrderedChildTerms(dc, childterms);
            if (dcchildren.length > 0) {
                result.push(...dcchildren);
            }
        });
        return result;
    }

    @trace(TraceLevel.Internal)
    protected async getItemById_Internal(id: string): Promise<T> {
        let result = null;
        await this.Init();
        const allTerms = await this.getAll_Query();
        if (allTerms && allTerms.length > 0) {
            const items = this.populateTerms(allTerms);
            //find and populate item
            result = find(items, (t) => t.id == id);
            if (!result) {
                console.warn(`[${this.serviceName}] - term with id ${id} not found`);
            }
        }
        return result;
    }

    @trace(TraceLevel.Internal)
    protected async getItemsById_Internal(
        ids: Array<string>
    ): Promise<Array<T>> {
        const results = new Array<T>();
        await this.Init();
        const allTerms = await this.getAll_Query();
        if (allTerms && allTerms.length > 0) {
            const items = this.populateTerms(allTerms);
            for (const id of ids) {
                const temp = find(items, (t) => t.id == id);
                if (temp) {
                    results.push(temp);
                } else {
                    console.warn(`[${this.serviceName}] - term with id ${id} not found`);
                }
            }
        }
        return results;
    }

    protected async get_Query(query: any): Promise<Array<any>> { // eslint-disable-line @typescript-eslint/no-unused-vars
        throw new Error("Not Implemented");
    }

    public async getItemById_Query(id: string): Promise<any> {
        console.error("[" + this.serviceName + ".getItemById_Query] - " + id);
        throw new Error("Not implemented");
    }
    public async getItemsById_Query(ids: Array<string>): Promise<Array<any>> {
        console.error(
            "[" + this.serviceName + ".getItemsById_Query] - " + ids.join(", ")
        );
        throw new Error("Not implemented");
    }

    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        console.error(
            "[" +
            this.serviceName +
            ".addOrUpdateItem_Internal] - " +
            JSON.stringify(item)
        );
        throw new Error("Not implemented");
    }

    protected async addOrUpdateItems_Internal(
        items: Array<T> /*, onItemUpdated?: (oldItem: T, newItem: T) => void*/
    ): Promise<Array<T>> {
        console.error(
            "[" +
            this.serviceName +
            ".addOrUpdateItems_Internal] - " +
            JSON.stringify(items)
        );
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: T): Promise<T> {
        console.error(
            "[" + this.serviceName + ".deleteItem_Internal] - " + JSON.stringify(item)
        );
        throw new Error("Not implemented");
    }

    protected async deleteItems_Internal(items: Array<T>): Promise<Array<T>> {
        console.error(
            "[" +
            this.serviceName +
            ".deleteItems_Internal] - " +
            JSON.stringify(items)
        );
        throw new Error("Not implemented");
    }

    protected async recycleItem_Internal(item: T): Promise<T> {
        console.error(
            "[" + this.serviceName + ".recycleItem_Internal] - " + JSON.stringify(item)
        );
        throw new Error("Not implemented");
    }

    protected async recycleItems_Internal(items: Array<T>): Promise<Array<T>> {
        console.error(
            "[" +
            this.serviceName +
            ".recycleItems_Internal] - " +
            JSON.stringify(items)
        );
        throw new Error("Not implemented");
    }
}
