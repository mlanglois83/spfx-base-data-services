import { find } from "@microsoft/sp-lodash-subset";
import { PnPClientStorage } from "@pnp/common";
import { dateAdd, stringIsNullOrEmpty } from "@pnp/common/util";
import { sp } from "../../pnpExtensions/TermsetExt";
import "@pnp/sp/sites";
import { IOrderedTermInfo, ITermSet } from "@pnp/sp/taxonomy";
import "@pnp/sp/webs";
import { ServicesConfiguration } from "../../configuration";
import { TraceLevel } from "../../constants/index";
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
export class BaseTermsetService<T extends TaxonomyTerm> extends BaseDataService<T> {

    protected utilsService: UtilsService;
    protected termsetnameorid: string;
    protected isGlobal: boolean;
    protected static siteCollectionTermsetId: string;

    /**
     * Get site collection group
     */
    protected static _siteCollectionGroupIdPromise: Promise<string> = undefined;
    protected static _siteCollectionGroupId: string = undefined;
    protected static async getSiteCollectionGroupId(): Promise<string> {
        if(stringIsNullOrEmpty(BaseTermsetService._siteCollectionGroupId)) {
            if (!BaseTermsetService._siteCollectionGroupIdPromise) {
                BaseTermsetService._siteCollectionGroupIdPromise = new Promise<string>(async (resolve, reject) => {
                    try {
                        const [ts, properties] = await Promise.all([
                            sp.termStore.get(),
                            sp.site.rootWeb.allProperties.get()
                        ]);
                        BaseTermsetService._siteCollectionGroupId = properties["SiteCollectionGroupId" + ts.id];
                        resolve(BaseTermsetService._siteCollectionGroupId);
                    } catch (error) {
                        reject(error);
                    }
                });
                BaseTermsetService._siteCollectionGroupIdPromise.then(() => {
                    BaseTermsetService._siteCollectionGroupIdPromise = undefined;
                }).catch(() => {
                    BaseTermsetService._siteCollectionGroupIdPromise = undefined;
                });
            }
            return BaseTermsetService._siteCollectionGroupIdPromise; 
        }
        else {
            return BaseTermsetService._siteCollectionGroupId;
        }
    }

    protected static _termStoreLanguagePromise: Promise<string> = undefined;
    protected static _termStoreDefaultLanguageTag: string = undefined;
    protected static async initTermStoreDefaultLanguageTag(): Promise<void> {
        if(stringIsNullOrEmpty(BaseTermsetService._termStoreDefaultLanguageTag)) {
            if (!BaseTermsetService._termStoreLanguagePromise) {
                BaseTermsetService._termStoreLanguagePromise = new Promise<string>(async (resolve, reject) => {
                    try {

                        const ts = await sp.termStore.get();
                        BaseTermsetService._termStoreDefaultLanguageTag = ts.defaultLanguageTag;
                        resolve();
                    } catch (error) {
                        reject(error);
                    }
                });
                BaseTermsetService._termStoreLanguagePromise.then(() => {
                    BaseTermsetService._termStoreLanguagePromise = undefined;
                }).catch(() => {
                    BaseTermsetService._termStoreLanguagePromise = undefined;
                });
            }       
        }
    }


    /**
     * Associated termset (pnpjs)
     */
    protected _tsIdPromise  = undefined;
    protected _tsId = undefined;    
    @trace(TraceLevel.ServiceUtilities)
    protected async GetTermset(): Promise<ITermSet> {
        if(stringIsNullOrEmpty(this._tsId) && this.termsetnameorid.match(/[A-z0-9]{8}-([A-z0-9]{4}-){3}[A-z0-9]{12}/)) {
            this._tsId = this.termsetnameorid;
        }
        if(stringIsNullOrEmpty(this._tsId)) {
            if (!this._tsIdPromise) {
                this._tsIdPromise = new Promise<ITermSet>(async (resolve, reject) => {
                    try {
                        
                        if (this.isGlobal) {
                            const [termsets] = await Promise.all([
                                sp.termStore.sets.get(),
                                BaseTermsetService.initTermStoreDefaultLanguageTag()
                            ]);
                            const ts = find(termsets, t => t.localizedNames.some(ln => ln.languageTag === BaseTermsetService._termStoreDefaultLanguageTag && ln.name === this.termsetnameorid));
                            if(ts) {
                                this._tsId = ts.id;
                                resolve(sp.termStore.sets.getById(this._tsId));
                            }
                            else {
                                reject(new Error("Termset not found: " + this.termsetnameorid));
                            }
                        }
                        else {
                            const groupId = await BaseTermsetService.getSiteCollectionGroupId();
                            const [termsets] = await Promise.all([
                                sp.termStore.groups.getById(groupId).sets.get(),
                                BaseTermsetService.initTermStoreDefaultLanguageTag()
                            ]);
                            const ts = find(termsets, t => t.localizedNames.some(ln => ln.languageTag === BaseTermsetService._termStoreDefaultLanguageTag && ln.name === this.termsetnameorid));
                            if(ts) {
                                this._tsId = ts.id;
                                resolve(sp.termStore.sets.getById(this._tsId));
                            }
                            else {
                                reject(new Error("Termset not found in site collection group: " + this.termsetnameorid));
                            }
                        }
                    } catch (error) {
                        reject(error);
                    }
                });
                this._tsIdPromise.then(() => {
                    this._tsIdPromise = undefined;
                }).catch(() => {
                    this._tsIdPromise = undefined;
                });
            }
            return this._tsIdPromise;            
        }
        else {
            return sp.termStore.sets.getById(this._tsId);
        }
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
        const [taxonomyHiddenItems] = await Promise.all([
            ServiceFactory.getService(TaxonomyHidden).getAll(),
            BaseTermsetService.initTermStoreDefaultLanguageTag()
        ]);
        this.updateInitValues(TaxonomyHidden["name"], ...taxonomyHiddenItems);
    }

    protected populateTerms(terms: IOrderedTermInfo[], basePath?: string): Array<T> {
        const result = new Array<T>();
        for (const term of terms) {
            const item = this.populateTerm(term, basePath);
            result.push(item);
            if(term.childrenCount > 0) {               
                result.push(...this.populateTerms(term.children, item.path));
            }
        }        
        return result;
    }

    protected populateTerm(term: IOrderedTermInfo, basePath: string): T {
        const result: T = new this.itemType();
        // common properties
        result.id = term.id;
        result.isDeprecated = term.isDeprecated;
        // translated
        result.title = this.getTermTitle(term);
        result.description = this.getTermDescription(term);
        // path
        result.path = stringIsNullOrEmpty(basePath) ? term.defaultLabel : (basePath + ";" + term.defaultLabel);
        // properties
        result.customProperties = this.getTermProperties(term);
        // wssids
        const taxonomyHiddenItems = this.getServiceInitValues(TaxonomyHidden);
        result.wssids = [];
        for (const taxonomyHiddenItem of taxonomyHiddenItems) {
            if (taxonomyHiddenItem.termId == result.id) { result.wssids.push(taxonomyHiddenItem.id); }
        }
        return result;
    }

    protected getTermTitle(term: IOrderedTermInfo): string {
        return  this.getTranslatedLabel(term.labels.filter(l => l.isDefault).map(l => {return {label: l.name, languageTag: l.languageTag};}));
    }

    protected getTermDescription(term: IOrderedTermInfo): string {
        return this.getTranslatedLabel(term.descriptions.map(l => {return {label: l.description, languageTag: l.languageTag};}));
    }

    protected getTermPath(term: IOrderedTermInfo, allTerms: IOrderedTermInfo[]): string {
        const parts = [this.getTermTitle(term)];
        while(term.parent) {
            term = find(allTerms, t => t.id === term.parent.id);
            parts.push(this.getTermTitle(term));
        }
        return parts.reverse().join(";");
    }
    protected getTermProperties(term: IOrderedTermInfo): {[key: string]: string} {
        const result: {[key: string]: string} = {};
        if(term.properties) {
            term.properties.forEach(p => {
                result[p.key] = p.value;
            });
        }
        return result;
    }

    protected getTranslatedLabel(labelCollection: {label: string; languageTag: string}[]): string {
        // current ui language
        const current = ServicesConfiguration.context.pageContext.cultureInfo.currentUICultureName;
        const currentLabel = find(labelCollection, label => label.languageTag === current);
        if(currentLabel) {
            return currentLabel.label;
        }        
        else {   
            // web language         
            const web = ServicesConfiguration.context.pageContext.web.languageName;
            const webLabel = find(labelCollection, label => label.languageTag === web);
            if(webLabel) {
                return webLabel.label;
            }
            else {      
                // default termstore language          
                const taxonomy = BaseTermsetService._termStoreDefaultLanguageTag;                
                const taxonomyLabel = find(labelCollection, label => label.languageTag === taxonomy);
                if(webLabel) {
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
        return taxonomyHiddenItems.filter((taxItem) => {
            return taxItem.termId === termId;
        }).map((filteredItem) => {
            return filteredItem.id;
        });
    }
    @trace(TraceLevel.Queries)
    protected async getAll_Query(): Promise<Array<IOrderedTermInfo>> {
        const termset = await this.GetTermset(); 
        const store = new PnPClientStorage();
        return store.session.getOrPut(this.serviceName + "-alltermsordered", () => {
            return termset.getAllChildrenAsOrderedTreeFull();
        }, dateAdd(new Date(), "minute", this.cacheDuration));
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
            result = find(items, t => t.id == id);
            if(!result) {               
                console.warn(`[${this.serviceName}] - term with id ${id} not found`);
            }         
        }
        return result;
    }

    @trace(TraceLevel.Internal)
    protected async getItemsById_Internal(ids: Array<number | string>): Promise<Array<T>> {

        const results = new Array<T>();
        const allTerms = await this.getAll_Query();
        if (allTerms && allTerms.length > 0) {
            if (!this.initialized) {
                await this.Init();
            }
            const items = this.populateTerms(allTerms);
            for (const id of ids) {
                const temp = find(items, t => t.id == id);
                if(temp) {                    
                    results.push(temp); 
                }  
                else {
                    console.warn(`[${this.serviceName}] - term with id ${id} not found`);
                }
            }
        }
        return results;
    }

    protected async get_Query(query: any): Promise<Array<any>> {// eslint-disable-line @typescript-eslint/no-unused-vars
        throw new Error('Not Implemented');
    }

    public async getItemById_Query(id: string): Promise<any> {
        console.error("[" + this.serviceName + ".getItemById_Query] - " + id);
        throw new Error("Not implemented");
    }
    public async getItemsById_Query(ids: Array<string>): Promise<Array<any>> { 
        console.error("[" + this.serviceName + ".getItemsById_Query] - " + ids.join(", "));
        throw new Error("Not implemented");
    }

    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        console.error("[" + this.serviceName + ".addOrUpdateItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async addOrUpdateItems_Internal(items: Array<T>/*, onItemUpdated?: (oldItem: T, newItem: T) => void*/): Promise<Array<T>> {
        console.error("[" + this.serviceName + ".addOrUpdateItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }

    protected async deleteItem_Internal(item: T): Promise<T> {
        console.error("[" + this.serviceName + ".deleteItem_Internal] - " + JSON.stringify(item));
        throw new Error("Not implemented");
    }

    protected async deleteItems_Internal(items: Array<T>): Promise<Array<T>> {
        console.error("[" + this.serviceName + ".deleteItems_Internal] - " + JSON.stringify(items));
        throw new Error("Not implemented");
    }   
    
}
