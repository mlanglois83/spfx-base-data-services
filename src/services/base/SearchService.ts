import {
  ISearchResult, SearchQueryBuilder,
  SearchResults,
  SortDirection
} from "@pnp/sp/search";
import { assign, cloneDeep, find } from "lodash";
import { BaseSPService, ServiceFactory } from "..";
import { FieldType, LogicalOperator, QueryToken, TestOperator } from "../../constants";
import { IFieldDescriptor, ILogicalSequence, IPredicate, IQuery } from "../../interfaces";
import { BaseItem, SPItem, TaxonomyTerm } from "../../models";

/**
 *
 * Base service search
 */
export class SearchService<TKey extends string | number, T extends BaseItem<TKey>> extends BaseSPService<T> {  
  protected recycleItem_Internal(item: T): Promise<T> {
    throw new Error("Method not implemented." + item.toString());
  }
  protected recycleItems_Internal(items: T[]): Promise<T[]> {
    throw new Error("Method not implemented." + items.toString());
  }
  protected deleteItem_Internal(item: T): Promise<T> {
    throw new Error("Method not implemented." + item.toString());
  }
  protected getAll_Query(linkedFields?: string[]): Promise<any[]> {
    throw new Error("Method not implemented." + linkedFields.toString());
  }
  protected get_Query(query: IQuery<T>, linkedFields?: string[]): Promise<any[]> {
    throw new Error("Method not implemented." + query.toString() + linkedFields.toString());
  }
  protected getItemById_Query(id: string | number, linkedFields?: string[]): Promise<any> {
    throw new Error("Method not implemented." + id.toString() + linkedFields.toString());
  }
  protected getItemsById_Query(id: (string | number)[], linkedFields?: string[]): Promise<any> {
    throw new Error("Method not implemented." + id.toString() + linkedFields.toString());
  }
  protected deleteItems_Internal(items: T[]): Promise<T[]> {
    throw new Error("Method not implemented." + items.toString());
  }


  protected async getAll_Internal(linkedFields?: Array<string>): Promise<Array<T>> {
    throw new Error("Not applicable" + linkedFields.toString());
  }

  protected async getItemById_Internal(id: number, linkedFields?: Array<string>): Promise<T> {
    throw new Error("Not implement yet" + id.toString() + linkedFields.toString());
  }
  protected async getItemsById_Internal(ids: Array<number>, linkedFields?: Array<string>): Promise<Array<T>> {
    throw new Error("Not implement yet" + ids.toString() + linkedFields.toString());
  }


  protected async addOrUpdateItem_Internal(item: T): Promise<T> {
    throw new Error("Not applicable" + item.toString());
  }

  protected async addOrUpdateItems_Internal(items: Array<T>, onItemUpdated?: (oldItem: T, newItem: T) => void): Promise<Array<T>> {
    throw new Error("Not applicable" + items.toString() + onItemUpdated.toString());
  }

  constructor(
    type: new (item?: any) => T,
    cacheDuration?: number,
    baseUrl?: string
  ) {
    super(type, cacheDuration, baseUrl);
  }

  protected _itemfields: any = null;
  protected _selectedProperties: Array<string> = [];

  protected initValues: any = {};

  private initPromise: Promise<void> = null;



  ///to validate all model dependencies (taxo) is loaded
  protected get isInitialized(): boolean {
    return this.initialized;
  }

  //properties to load in search query
  public get SelectedProperties(): Array<string> {
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

      Object.keys(result).forEach(propertyName => {
        const fieldDescription = result[propertyName];
        if (fieldDescription.fieldName) {
          this._selectedProperties.push(fieldDescription.fieldName);
        }
      });

      this._itemfields = result;
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
        } else {
          this.initValues = {};
          try {
            const fields = this.ItemFields;
            const models = [];
            for (const key in fields) {
              if (fields.hasOwnProperty(key)) {
                const fieldDescription = fields[key];
                // REM MLS : lookup removed from preload
                if (
                  fieldDescription.modelName &&
                  models.indexOf(fieldDescription.modelName) === -1 &&
                  fieldDescription.fieldType !== FieldType.Lookup &&
                  fieldDescription.fieldType !== FieldType.LookupMulti
                ) {
                  models.push(fieldDescription.modelName);
                }
              }
            }
            await Promise.all(
              models.map(async modelName => {
                if (!this.initValues[modelName]) {
                  const service = ServiceFactory.getServiceByModelName(modelName);
                  const values = await service.getAll();
                  this.initValues[modelName] = values;
                }
              })
            );
            this.initialized = true;
            this.initPromise = null;
            resolve();
          } catch (error) {
            this.initPromise = null;
            reject(error);
          }
        }
      });
    }
    return this.initPromise;
  }


  protected async get_Internal(query: IQuery<T>, linkedFields?: Array<string>): Promise<Array<T>> {

    console.log(linkedFields);

    await this.LoadTaxonomyDependency();
    //Generate query from intermediate language
    let builder = SearchQueryBuilder(
      query.test.type === "predicate"
        ? this.getPredicate(query.test)
        : this.getLogicalSequence(query.test)
    );
    builder = builder.selectProperties(...this.SelectedProperties);

    //manage order by from query
    if (query.orderBy) {
      const sorts = [];

      for (const sort of query.orderBy) {

        const fields = this.ItemFields;
        const field = fields[sort.propertyName];


        sorts.push({
          Property: field.fieldName,
          Direction: sort.ascending
            ? SortDirection.Ascending
            : SortDirection.Descending
        });
      }
      builder = builder.sortList(...sorts);
    }

    //mange limite
    if (query.limit) {
      builder = builder.rowLimit(query.limit);
    }

    // Start row
    if (query.lastId) {
      builder = builder.startRow(parseInt(query.lastId.toString()));
    }

    const searchQuery = builder.toSearchQuery();
    searchQuery.TrimDuplicates = false;

    //Execute query
    const searchItems: SearchResults = await this.sp.search(searchQuery);

    //browse results
    const results = searchItems.PrimarySearchResults.map(r => {
      //convert data
      return this.getItemFromSearchResult(r);
    });

    return results;
  }

  public async get_AllWithTotalRows(
    query: IQuery<T>
  ): Promise<{ searchResults: SearchResults; items: Array<T> }> {
    await this.LoadTaxonomyDependency();
    //Generate query from intermediate language
    let builder = SearchQueryBuilder(
      query.test.type === "predicate"
        ? this.getPredicate(query.test)
        : this.getLogicalSequence(query.test)
    );
    builder = builder.selectProperties(...this.SelectedProperties);

    //manage order by from query
    if (query.orderBy) {
      const sorts = [];

      for (const sort of query.orderBy) {
        sorts.push({
          Property: sort.propertyName,
          Direction: sort.ascending
            ? SortDirection.Ascending
            : SortDirection.Descending
        });
      }
      builder = builder.sortList(...sorts);
    }

    //mange limite
    if (query.limit) {
      builder = builder.rowLimit(query.limit);
    }

    // Start row
    if (query.lastId) {
      builder = builder.startRow(parseInt(query.lastId.toString()));
    }

    const searchQuery = builder.toSearchQuery();
    searchQuery.TrimDuplicates = false;
    //Execute query
    const searchItems: SearchResults = await this.sp.search(searchQuery);

    //browse results
    const results = searchItems.PrimarySearchResults.map(r => {
      //convert data
      return this.getItemFromSearchResult(r);
    });

    return { searchResults: searchItems, items: results };
  }

  /**
   * convert search result to object model
   * @param searchItem
   */
  private getItemFromSearchResult(searchItem: ISearchResult): T {
    const item = new this.itemType();

    //for each properties decorated
    Object.keys(this.ItemFields).forEach(propertyName => {
      const fieldDescription = this.ItemFields[propertyName];
      //get value in search result to assign to object model
      this.setFieldValue(searchItem, item, propertyName, fieldDescription);
    });
    return item;
  }



  /**
   *
   * @param termID - termID of term to retrieve
   * @param terms - terms list where term must be found
   */
  public getTaxonomyTermById<TermType extends TaxonomyTerm>(
    termId: string,
    terms: Array<TermType>
  ): TermType {
    return find(terms, term => {
      return term.id && term.id.indexOf(termId) > -1;
    });
  }

  private setFieldValue(
    searchItem: any,
    destItem: T,
    propertyName: string,
    fieldDescriptor: IFieldDescriptor
  ): void {

    const converted = (destItem as unknown) as SPItem;
    fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;

    switch (fieldDescriptor.fieldType) {
      case FieldType.Simple:
        converted[propertyName] = searchItem[fieldDescriptor.fieldName]
          ? searchItem[fieldDescriptor.fieldName]
          : fieldDescriptor.defaultValue;
        break;
      case FieldType.Boolean:
        converted[propertyName] = searchItem[fieldDescriptor.fieldName]
          ? searchItem[fieldDescriptor.fieldName] === "true"
          : fieldDescriptor.defaultValue;
        break;
      case FieldType.Number:
        converted[propertyName] = searchItem[fieldDescriptor.fieldName]
          ? Number(searchItem[fieldDescriptor.fieldName])
          : fieldDescriptor.defaultValue;
        break;
      case FieldType.Date:
        converted[propertyName] = searchItem[fieldDescriptor.fieldName]
          ? new Date(searchItem[fieldDescriptor.fieldName])
          : fieldDescriptor.defaultValue;
        break;
      case FieldType.Lookup:
        console.error(
          "lookup " + fieldDescriptor.fieldName + " not yet managed"
        );
        break;
      case FieldType.LookupMulti:
        console.error(
          "lookup " + fieldDescriptor.fieldName + " not yet managed"
        );
        break;
      case FieldType.User:
        console.error(
          "user " + fieldDescriptor.fieldName + " not yet managed"
        );

        break;
      case FieldType.UserMulti:
        console.error(
          "user " + fieldDescriptor.fieldName + " not yet managed"
        );
        break;
      case FieldType.Taxonomy:
        let termId;
        if (searchItem[fieldDescriptor.fieldName] && searchItem[fieldDescriptor.fieldName].includes("GP0|#")) {
          termId = searchItem[fieldDescriptor.fieldName].split("\n").filter((str) => { return str.indexOf("GP0|#") === 0; }).map((str) => { return str.replace("GP0|#", ""); })[0];
        }
        else {
          termId = searchItem[fieldDescriptor.fieldName]
            ? searchItem[fieldDescriptor.fieldName].split("|")[1]
            : null;
        }

        if (termId) {
          const tterms = this.getServiceInitValuesByName(fieldDescriptor.modelName);
          const retrievedTerm  = this.getTaxonomyTermById(termId, tterms as Array<TaxonomyTerm>);
          converted[propertyName] = retrievedTerm ?? fieldDescriptor.defaultValue;
        }
        else {
          converted[propertyName] = fieldDescriptor.defaultValue;
        }
        break;
      case FieldType.TaxonomyMulti:
        let terms;
        if (searchItem[fieldDescriptor.fieldName] && searchItem[fieldDescriptor.fieldName].includes("GP0|#")) {
          terms = searchItem[fieldDescriptor.fieldName].split(";").filter((str) => { return str.indexOf("GP0|#") === 0; }).map((str) => { return str.replace("GP0|#", ""); });
        }
        else {
          terms = searchItem[fieldDescriptor.fieldName]
            ? searchItem[fieldDescriptor.fieldName].split(";")
            : null;
        }

        if (terms) {
          converted[propertyName] = [];
          const tterms = this.getServiceInitValuesByName(fieldDescriptor.modelName);
          terms.forEach(term => {
            const tempId: string = term ? term.split("|")[1] : null;
            if (tempId) {
              const retrievedTerm = this.getTaxonomyTermById(tempId, tterms as Array<TaxonomyTerm>);
              if(retrievedTerm) {                
                converted[propertyName].push(retrievedTerm);
              }
            }
          });
        }
        else {
          converted[propertyName] = fieldDescriptor.defaultValue;
        }
        break;
      case FieldType.Json:
        converted[propertyName] = searchItem[fieldDescriptor.fieldName]
          ? JSON.parse(searchItem[fieldDescriptor.fieldName])
          : fieldDescriptor.defaultValue;
        break;
    }
  }

  private getLogicalSequence(sequence: ILogicalSequence<T>): string {
    const cloneSequence = cloneDeep(sequence);
    if (!cloneSequence.children || cloneSequence.children.length === 0) {
      return "";
    }
    // if (cloneSequence.children.length === 1)
    //        // first part

    let result = "";

    cloneSequence.children.length === 1 ? result += "" : result += "(";

    const subQueries = cloneSequence.children.map(subSequence => {
      if (subSequence.type === "predicate") {
        return this.getPredicate(subSequence as IPredicate<T, keyof T>);
      } else {
        return this.getLogicalSequence(subSequence as ILogicalSequence<T>);
      }
    });

    let firstOrEmptySubQuery = true;
    for (let _i = 0; _i < subQueries.length; _i++) {

      if (subQueries[_i] && subQueries[_i] != "") {

        firstOrEmptySubQuery ? result += subQueries[_i] : result += cloneSequence.operator == LogicalOperator.And ? " AND " + subQueries[_i] : " OR " + subQueries[_i];

        firstOrEmptySubQuery = false;
      }

    }

    cloneSequence.children.length === 1 ? result += "" : result += ")";

    return result;

  }

  private getPredicate(predicate: IPredicate<T, keyof T>): string {
    let result = "";

    let valueTransformForRequest = "";

    if (predicate.operator !== TestOperator.FreeRequest && predicate.operator !== TestOperator.IsNotNull && predicate.operator !== TestOperator.IsNull) {
      valueTransformForRequest = this.getValue(predicate);
    }

    const fields = this.ItemFields;
    const field = fields[predicate.propertyName.toString()];


    switch (predicate.operator) {
      case TestOperator.BeginsWith:
        result = valueTransformForRequest;
        break;
      case TestOperator.Eq:
        result = field.fieldName + ":" + valueTransformForRequest;
        break;

      case TestOperator.Neq:
        result = "-" + field.fieldName + ":" + valueTransformForRequest;
        break;
      case TestOperator.Contains:
        result = field.fieldName + ":" + valueTransformForRequest + "*";
        break;
      case TestOperator.Geq:
        result = field.fieldName + ">=" + valueTransformForRequest;
        break;
      case TestOperator.Leq:
        result = field.fieldName + "<=" + valueTransformForRequest;
        break;
      case TestOperator.FreeRequest:
        result = predicate.value;

        break;

      case TestOperator.IsNotNull:
        result = " (Field:a* OR Field:b* OR Field:c* OR Field:d* OR Field:e* OR Field:f* OR Field:g* OR Field:h* OR Field:i* OR Field:j* OR Field:k* OR Field:l* OR Field:m* OR Field:n* OR Field:o* OR Field:p* OR Field:q* OR Field:r* OR Field:s* OR Field:t* OR Field:u* OR Field:v* OR Field:w* OR Field:x* OR Field:y* OR Field:z* OR Field:1* OR Field:2* OR Field:3* OR Field:4* OR Field:5* OR Field:6* OR Field:7* OR Field:8* OR Field:9* OR Field:0*) ".replace(
          /Field/gi,
          field.fieldName.toString()
        );

        break;
      default:
        throw new Error(
          "Operator " + predicate.operator + " not yet implemented."
        );
    }
    return result;
  }

  private getValue(predicate: IPredicate<T, keyof T>): string {
    let transformValue = "";

    const fields = this.ItemFields;
    const field = fields[predicate.propertyName.toString()];


    const { propertyName, value } = predicate;
    if (field) {
      switch (field.fieldType) {
        case FieldType.Date:
          if (value === QueryToken.Now || value === QueryToken.Today) {
            transformValue = new Date().toISOString();
          } else {
            if (value == null) {
              transformValue = null;
            }
            else {
              transformValue = value.toISOString();
            }
          }
          break;
        case FieldType.Json:
          console.error(
            "type value Json not yet implement. " + field.fieldName
          );
          break;
        case FieldType.Lookup:
        case FieldType.LookupMulti:
          console.error(
            "type value Lookup not yet implement. " + field.fieldName
          );
          break;
        case FieldType.Taxonomy:
        case FieldType.TaxonomyMulti:
          transformValue = (value as TaxonomyTerm).id;
          break;
        case FieldType.User:
        case FieldType.UserMulti:
          console.error(
            "type value user not yet implement. " + field.fieldName
          );
          break;
        case FieldType.Simple:

        default:
          if (typeof value === "number") {
            transformValue = value.toString();
          } else if (typeof value === "boolean") {
            console.error(
              "type value boolean not yet implement. " + field.fieldName
            );
            transformValue = value ? "1" : "0";
          } else {
            transformValue = value.toString();
          }
          break;
      }
    } else {
      throw new Error("Field was not found : " + propertyName.toString());
    }

    return transformValue;
  }
}
