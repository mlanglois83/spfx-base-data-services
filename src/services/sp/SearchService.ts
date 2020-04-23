
import { assign, cloneDeep, find } from "@microsoft/sp-lodash-subset";
import { SearchQueryBuilder, SearchResults, SortDirection, ISearchResult } from "@pnp/sp/search";
import { sp } from "@pnp/sp";

import { IBaseItem, BaseDataService, FieldType, SPItem, TaxonomyTerm, IFieldDescriptor, LogicalOperator, IPredicate, ILogicalSequence, TestOperator, IQuery, ServicesConfiguration } from "../..";

/**
 * 
 * Base service search
 */
export class SearchService<T extends IBaseItem> extends BaseDataService<T>{

    constructor(type: (new (item?: any) => T), tableName: string, cacheDuration?: number) {
        super(type, tableName, cacheDuration);

    }

    protected _itemfields: any = null;
    protected _selectedProperties: Array<string> = [];

    protected initValues: any = {};

    private initPromise: Promise<void> = null;

    private initialized = false;


    ///to validate all model dependencies (taxo) is loaded
    protected get isInitialized(): boolean {
        return this.initialized;
    }

    //properties to load in search query
    public get SelectedProperties(): Array<string> {
        const temp = this.ItemFields;
        console.log(temp);
        return this._selectedProperties;
    }
    /***
     * get Fields and their configuration (decorator) from model
     */
    public get ItemFields(): any {
        let result = {};

        if (!this._itemfields) {
            assign(result, this.itemType["Fields"][SPItem["name"]]);
            if (this.itemType["Fields"][this.itemType["name"]]) {
                assign(result, this.itemType["Fields"][this.itemType["name"]]);
            }

            Object.keys(result).map((propertyName) => {
                const fieldDescription = result[propertyName];
                if (fieldDescription.searchName) {
                    this._selectedProperties.push(fieldDescription.searchName);
                }
            });

            this._itemfields = result;
            console.log(this._itemfields);
            console.log(this._selectedProperties);
        }

        result = this._itemfields;
        return result;
    }


    /**
     * Load model taxonomy dependencies
     * 
     */
    public async LoadTaxonomyDependency(): Promise<void> {
        if (!this.initPromise) {
            this.initPromise = new Promise<void>(async (resolve, reject) => {
                if (this.initialized) {
                    resolve();
                }
                else {
                    this.initValues = {};
                    try {

                        const fields = this.ItemFields;
                        const models = [];
                        for (const key in fields) {
                            if (fields.hasOwnProperty(key)) {
                                const fieldDescription = fields[key];
                                // REM MLS : lookup removed from preload
                                if (fieldDescription.modelName && models.indexOf(fieldDescription.modelName) === -1 &&
                                    fieldDescription.fieldType !== FieldType.Lookup &&
                                    fieldDescription.fieldType !== FieldType.LookupMulti) {
                                    models.push(fieldDescription.modelName);
                                }
                            }
                        }
                        await Promise.all(models.map(async (modelName) => {
                            if (!this.initValues[modelName]) {
                                const service = ServicesConfiguration.configuration.serviceFactory.create(modelName);
                                const values = await service.getAll();
                                this.initValues[modelName] = values;
                            }
                        }));
                        this.initialized = true;
                        this.initPromise = null;
                        resolve();
                    }
                    catch (error) {
                        this.initPromise = null;
                        reject(error);
                    }
                }
            });
        }
        return this.initPromise;

    }

    protected async addOrUpdateItem_Internal(item: T): Promise<T> {
        console.log(item);
        throw new Error("Not applicable");
    }

    protected async addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
        console.log(items);
        throw new Error("Not applicable");
    }

    protected async deleteItem_Internal(item: T): Promise<void> {
        console.log(item);
        throw new Error("Not applicable");
    }

    protected async getAll_Internal(linkedFields?: Array<string>): Promise<Array<T>> {
        console.log(linkedFields);
        throw new Error("Not applicable");
    }

    protected async getItemById_Internal(id: number, linkedFields?: Array<string>): Promise<T> {
        console.log(id);
        console.log(linkedFields);
        throw new Error("Not implement yet");

    }
    protected async getItemsById_Internal(ids: Array<number>, linkedFields?: Array<string>): Promise<Array<T>> {
        console.log(ids);
        console.log(linkedFields);
        throw new Error("Not implement yet");
    }



    protected async get_Internal(query: IQuery, linkedFields?: Array<string>): Promise<Array<T>> {
        console.log(linkedFields);
        await this.LoadTaxonomyDependency();
        //Generate query from intermediate language
        let builder = SearchQueryBuilder(
            query.test.type === "predicate" ? this.getPredicate(query.test) : this.getLogicalSequence(query.test)

        );
        builder = builder.selectProperties(...this.SelectedProperties);

        //manage order by from query
        if (query.orderBy) {
            const sorts = [];

            for (const sort of query.orderBy) {
                sorts.push({
                    Property: sort.propertyName,
                    Direction: sort.ascending ? SortDirection.Ascending : SortDirection.Descending
                });
            }
            builder = builder.sortList(...sorts);
        }

        //mange limite
        if (query.limit) {
            builder = builder.rowLimit(query.limit);
        }

        //Execute query
        const searchItems: SearchResults = await sp.search(builder);

        //browse results
        const results = searchItems.PrimarySearchResults.map((r) => {
            //convert data
            return this.getItemFromSearchResult(r);
        });

        return results;
    }

    /**
     * convert search result to object model
     * @param searchItem 
     */
    private getItemFromSearchResult(searchItem: ISearchResult): T {
        const item = new this.itemType();

        //for each properties decorated
        Object.keys(this.ItemFields).map((propertyName) => {
            const fieldDescription = this.ItemFields[propertyName];
            //get value in search result to assign to object model
            this.setFieldValue(searchItem, item, propertyName, fieldDescription);
        });
        return item;
    }

    private getServiceInitValues(modelName: string): any {
        return this.initValues[modelName];
    }

    /**
     * 
     * @param termID - termID of term to retrieve
     * @param terms - terms list where term must be found
     */
    public getTaxonomyTermById<TermType extends TaxonomyTerm>(termId: string, terms: Array<TermType>): TermType {
        return find(terms, (term) => {
            return (term.id && term.id.indexOf(termId) > -1);
        });
    }

    private setFieldValue(searchItem: any, destItem: T, propertyName: string, fieldDescriptor: IFieldDescriptor): void {
        const converted = destItem as unknown as SPItem;
        fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;
        switch (fieldDescriptor.fieldType) {
            case FieldType.Simple:
                converted[propertyName] = searchItem[fieldDescriptor.searchName] ? searchItem[fieldDescriptor.searchName] : fieldDescriptor.defaultValue;
                break;
            case FieldType.Date:
                converted[propertyName] = searchItem[fieldDescriptor.searchName] ? new Date(searchItem[fieldDescriptor.searchName]) : fieldDescriptor.defaultValue;
                break;
            case FieldType.Lookup:
                console.error("lookup " + fieldDescriptor.searchName + " not yet managed");
                break;
            case FieldType.LookupMulti:
                console.error("lookup " + fieldDescriptor.searchName + " not yet managed");
                break;
            case FieldType.User:
                console.error("user " + fieldDescriptor.searchName + " not yet managed");
                // user service, not working for search because use local site to search users
                // const upn: string = searchItem[fieldDescriptor.searchName] ? searchItem[fieldDescriptor.searchName].split("|")[0].trim() : null;

                // console.log(upn);
                // if (!stringIsNullOrEmpty(upn)) {
                //     // get values from init values
                //     const users = this.getServiceInitValues(fieldDescriptor.modelName);
                //     const existing = find(users, (user) => {
                //         return user.userPrincipalName === upn;
                //     });
                //     converted[propertyName] = existing ? existing : fieldDescriptor.defaultValue;
                // }
                // else {
                //     converted[propertyName] = fieldDescriptor.defaultValue;
                // }
                break;
            case FieldType.UserMulti:
                console.error("user " + fieldDescriptor.searchName + " not yet managed");
                break;
            case FieldType.Taxonomy:
                const termId: string = searchItem[fieldDescriptor.searchName] ? searchItem[fieldDescriptor.searchName].split('|')[1] : null;

                if (termId) {
                    const tterms = this.getServiceInitValues(fieldDescriptor.modelName);
                    converted[propertyName] = this.getTaxonomyTermById(termId, tterms);
                }
                break;
            case FieldType.TaxonomyMulti:

                const terms: Array<string> = searchItem[fieldDescriptor.searchName] ? searchItem[fieldDescriptor.searchName].split(';') : null;

                if (terms) {
                    converted[propertyName] = [];
                    terms.map((term) => {
                        const tempId: string = term ? term.split('|')[1] : null;
                        if (tempId) {
                            const tterms = this.getServiceInitValues(fieldDescriptor.modelName);
                            converted[propertyName].push(this.getTaxonomyTermById(tempId, tterms));
                        }
                    });
                }
                break;
            case FieldType.Json:
                converted[propertyName] = searchItem[fieldDescriptor.searchName] ? JSON.parse(searchItem[fieldDescriptor.searchName]) : fieldDescriptor.defaultValue;
                break;
        }
    }


    private getLogicalSequence(sequence: ILogicalSequence): string {
        const cloneSequence = cloneDeep(sequence);
        if (!cloneSequence.children || cloneSequence.children.length === 0) {
            return "";
        }
        if (cloneSequence.children.length === 1) {
            if (cloneSequence.children[0].type === "predicate") {
                return this.getPredicate(cloneSequence.children[0] as IPredicate);
            }
            else {
                return this.getLogicalSequence(cloneSequence.children[0] as ILogicalSequence);
            }
        }
        else {
            // first part
            let result = cloneSequence.operator == LogicalOperator.And ? " AND " : " OR " + "(";
            if (cloneSequence.children[0].type === "predicate") {
                result += this.getPredicate(cloneSequence.children[0] as IPredicate);
            }
            else {
                result += this.getLogicalSequence(cloneSequence.children[0] as ILogicalSequence);
            }
            cloneSequence.children.splice(0, 1);
            result += this.getLogicalSequence(cloneSequence);
            result += ")";
            return result;
        }
    }



    private getPredicate(predicate: IPredicate): string {
        let result = "";
        switch (predicate.operator) {
            case TestOperator.Eq:
                result = predicate.propertyName + ":" + predicate.value;
                break;

            case TestOperator.Neq:
                result = "-" + predicate.propertyName + ":" + predicate.value;
                break;

            case TestOperator.Contains:
                result = predicate.propertyName + ":" + predicate.value + "*";
                break;
            default:
                throw new Error("Operator " + predicate.operator + " nto yet implement.");
        }
        return result;
    }
}