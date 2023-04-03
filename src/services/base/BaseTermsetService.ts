import { find } from "lodash";
import { dateAdd, stringIsNullOrEmpty, PnPClientStorage } from "@pnp/core";
import "../../pnpExtensions/TermsetExt";
import "@pnp/sp/sites";
import { IOrderedTermInfo, ITermSet } from "@pnp/sp/taxonomy";
import "@pnp/sp/webs";
import { ServicesConfiguration } from "../../configuration";
import { Constants, TraceLevel } from "../../constants/index";
import { Decorators } from "../../decorators";
import { TaxonomyHidden, TaxonomyTerm } from "../../models";
import { ServiceFactory } from "../ServiceFactory";
import { UtilsService } from "../UtilsService";
import { BaseDataService } from "./BaseDataService";
const trace = Decorators.trace;
const standardTermSetCacheDuration = 10;

/**
 * Base service for sp termset operations
 */
export class BaseTermsetService<
    T extends TaxonomyTerm
    > extends BaseDataService<T> {
    protected utilsService: UtilsService;
    protected termsetnameorid: string;
    protected isGlobal: boolean;
    protected static siteCollectionTermsetId: string;

    
    protected set customSortOrder(value: string) {
        localStorage.setItem(UtilsService.formatText(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.serverRelativeUrl, this.serviceName), value ? value : "");
    }
    protected get customSortOrder(): string {
        return localStorage.getItem(UtilsService.formatText(Constants.cacheKeys.termsetCustomOrder, ServicesConfiguration.serverRelativeUrl, this.serviceName));
    }

    /**
     * Get site collection group
     */
    protected static _siteCollectionGroupIdPromise: Promise<string> = undefined;
    protected static _siteCollectionGroupId: string = undefined;
    protected static async getSiteCollectionGroupId(): Promise<string> {
        if (stringIsNullOrEmpty(BaseTermsetService._siteCollectionGroupId)) {
            if (!BaseTermsetService._siteCollectionGroupIdPromise) {
                BaseTermsetService._siteCollectionGroupIdPromise = new Promise<string>(
                    async (resolve, reject) => {
                        try {
                            const [ts, properties] = await Promise.all([
                                ServicesConfiguration.sp.termStore(),
                                ServicesConfiguration.sp.site.rootWeb.allProperties(),
                            ]);
                            BaseTermsetService._siteCollectionGroupId =
                                properties["SiteCollectionGroupId" + ts.id] ||
                                properties[
                                "SiteCollectionGroupId" + ts.id.replace(/-/g, "_x002d_")
                                ];
                            resolve(BaseTermsetService._siteCollectionGroupId);
                        } catch (error) {
                            reject(error);
                        }
                    }
                );
                BaseTermsetService._siteCollectionGroupIdPromise
                    .then(() => {
                        BaseTermsetService._siteCollectionGroupIdPromise = undefined;
                    })
                    .catch(() => {
                        BaseTermsetService._siteCollectionGroupIdPromise = undefined;
                    });
            }
            return BaseTermsetService._siteCollectionGroupIdPromise;
        } else {
            return BaseTermsetService._siteCollectionGroupId;
        }
    }

    protected static _termStoreLanguagePromise: Promise<string> = undefined;
    protected static _termStoreDefaultLanguageTag: string = undefined;
    protected static async initTermStoreDefaultLanguageTag(): Promise<string> {
        if (stringIsNullOrEmpty(BaseTermsetService._termStoreDefaultLanguageTag)) {
            if (!BaseTermsetService._termStoreLanguagePromise) {
                BaseTermsetService._termStoreLanguagePromise = new Promise<string>(
                    async (resolve, reject) => {
                        try {
                            const ts = await ServicesConfiguration.sp.termStore();
                            BaseTermsetService._termStoreDefaultLanguageTag =
                                ts.defaultLanguageTag;
                            resolve(ts.defaultLanguageTag);
                        } catch (error) {
                            reject(error);
                        }
                    }
                );
                BaseTermsetService._termStoreLanguagePromise
                    .then(() => {
                        BaseTermsetService._termStoreLanguagePromise = undefined;
                    })
                    .catch(() => {
                        BaseTermsetService._termStoreLanguagePromise = undefined;
                    });
            }
            return BaseTermsetService._termStoreLanguagePromise;
        }
        return BaseTermsetService._termStoreDefaultLanguageTag;
    }

    /**
     * Associated termset (pnpjs)
     */
    protected _tsIdPromise = undefined;
    protected _tsId = undefined;
    @trace(TraceLevel.ServiceUtilities)
    protected async GetTermset(): Promise<ITermSet> {
        if (
            stringIsNullOrEmpty(this._tsId) &&
            this.termsetnameorid.match(/[A-z0-9]{8}-([A-z0-9]{4}-){3}[A-z0-9]{12}/)
        ) {
            this._tsId = this.termsetnameorid;
        }
        if (stringIsNullOrEmpty(this._tsId)) {
            if (!this._tsIdPromise) {
                this._tsIdPromise = new Promise<ITermSet>(async (resolve, reject) => {
                    try {
                        if (this.isGlobal) {
                            const [termsets, tsLngTag] = await Promise.all([
                                ServicesConfiguration.sp.termStore.sets(),
                                BaseTermsetService.initTermStoreDefaultLanguageTag(),
                            ]);
                            const ts = find(termsets, (t) =>
                                t.localizedNames.some(
                                    (ln) =>
                                        ln.languageTag === tsLngTag &&
                                        ln.name === this.termsetnameorid
                                )
                            );
                            if (ts) {
                                this._tsId = ts.id;
                                this.customSortOrder = ts.customSortOrder?.join(":");
                                resolve(ServicesConfiguration.sp.termStore.sets.getById(this._tsId));
                            } else {
                                reject(new Error("Termset not found: " + this.termsetnameorid));
                            }
                        } else {
                            const groupId =
                                await BaseTermsetService.getSiteCollectionGroupId();
                            const [termsets, tsLngTag] = await Promise.all([
                                ServicesConfiguration.sp.termStore.groups.getById(groupId).sets(),
                                BaseTermsetService.initTermStoreDefaultLanguageTag(),
                            ]);
                            const ts = find(termsets, (t) =>
                                t.localizedNames.some(
                                    (ln) =>
                                        ln.languageTag === tsLngTag &&
                                        ln.name === this.termsetnameorid
                                )
                            );
                            if (ts) {
                                this._tsId = ts.id;
                                this.customSortOrder = ts.customSortOrder?.join(":");
                                resolve(ServicesConfiguration.sp.termStore.sets.getById(this._tsId));
                            } else {
                                reject(
                                    new Error(
                                        "Termset not found in site collection group: " +
                                        this.termsetnameorid
                                    )
                                );
                            }
                        }
                    } catch (error) {
                        reject(error);
                    }
                });
                this._tsIdPromise
                    .then(() => {
                        this._tsIdPromise = undefined;
                    })
                    .catch(() => {
                        this._tsIdPromise = undefined;
                    });
            }
            return this._tsIdPromise;
        } else {
            return ServicesConfiguration.sp.termStore.sets.getById(this._tsId);
        }
    }

    /**
     *
     * @param type - items type
     * @param context - current sp component context
     * @param termsetname - term set name
     */
    constructor(
        type: new (item?: any) => T,
        termsetnameorid: string,
        isGlobal = true,
        cacheDuration: number = standardTermSetCacheDuration
    ) {
        super(type, cacheDuration);
        this.utilsService = new UtilsService();
        this.termsetnameorid = termsetnameorid;
        this.isGlobal = isGlobal;
    }

    @trace(TraceLevel.ServiceUtilities)
    protected async init_internal(): Promise<void> {
        await super.init_internal();
        const [taxonomyHiddenItems] = await Promise.all([
            ServiceFactory.getService(TaxonomyHidden).getAll(),
            BaseTermsetService.initTermStoreDefaultLanguageTag(),
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
            const currentSortOrder = find(term.customSortOrder, cso => cso.setId === this._tsId);
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
                const taxonomy = BaseTermsetService._termStoreDefaultLanguageTag;
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
        if (!this.initialized) {
            await this.Init();
        }
        const taxonomyHiddenItems = this.getServiceInitValues(TaxonomyHidden);
        return taxonomyHiddenItems
            .filter((taxItem) => {
                return taxItem.termId === termId;
            })
            .map((filteredItem) => {
                return filteredItem.id;
            });
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
            dateAdd(new Date(), "minute", this.cacheDuration)
        );
    }

    @trace(TraceLevel.Internal)
    protected async getAll_Internal(): Promise<Array<T>> {
        let results: Array<T> = [];
        const items = await this.getAll_Query();
        if (items && items.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
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
        const allTerms = await this.getAll_Query();
        if (allTerms && allTerms.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
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
        ids: Array<number | string>
    ): Promise<Array<T>> {
        const results = new Array<T>();
        const allTerms = await this.getAll_Query();
        if (allTerms && allTerms.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
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
