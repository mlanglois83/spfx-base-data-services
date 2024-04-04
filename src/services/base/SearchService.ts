import { stringIsNullOrEmpty } from "@pnp/core";
import {
  SearchQueryBuilder,
  SearchResults,
  SortDirection
} from "@pnp/sp/search";
import { cloneDeep, find } from "lodash";
import { BaseSPService, ServiceFactory, UserService, UtilsService } from "..";
import { FieldType, LogicalOperator, QueryToken, TestOperator } from "../../constants";
import { IBaseSPServiceOptions, IFieldDescriptor, ILogicalSequence, IPredicate, IQuery } from "../../interfaces";
import { BaseItem, SPItem, TaxonomyTerm, User } from "../../models";

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
    itemType: (new (item?: any) => T),
    options?: IBaseSPServiceOptions, 
    ...args: any[]
  ) {
    super(itemType, options, ...args);
  }


  protected initValues: any = {};

  private initPromise: Promise<void> = null;

  ///to validate all model dependencies (taxo) is loaded
  protected get isInitialized(): boolean {
    return this.initialized;
  }

  //properties to load in search query
  public get SelectedProperties(): Array<string> {
    const result = [];
    const itemFields = super.ItemFields;
    if (itemFields) {      
      Object.keys(itemFields).forEach(propertyName => {
        const fieldDescription = itemFields[propertyName];
        if (fieldDescription.fieldName) {
          result.push(fieldDescription.fieldName);
        }
      });
    }
    return result;
  }

  
  protected async get_Query(query: IQuery<T>): Promise<any[]> {
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
        const field = fields[sort.propertyName.toString()];

        sorts.push({
          Property: field.fieldName,
          Direction: sort.ascending
            ? SortDirection.Ascending
            : SortDirection.Descending
        });
      }
      builder = builder.sortList(...sorts);
    }

    //manage limite
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
    return searchItems.PrimarySearchResults
  }
  


  protected populateFieldValue(data: any, destItem: T, propertyName: string, fieldDescriptor: IFieldDescriptor): void {
    const converted = (destItem as unknown) as SPItem;
    fieldDescriptor.fieldType = fieldDescriptor.fieldType || FieldType.Simple;

    switch (fieldDescriptor.fieldType) {
      case FieldType.Simple:
        converted[propertyName] = data[fieldDescriptor.fieldName]
          ? data[fieldDescriptor.fieldName]
          : fieldDescriptor.defaultValue;
        break;
      case FieldType.Boolean:
        converted[propertyName] = data[fieldDescriptor.fieldName]
          ? data[fieldDescriptor.fieldName] === "true"
          : fieldDescriptor.defaultValue;
        break;
      case FieldType.Number:
        converted[propertyName] = data[fieldDescriptor.fieldName]
          ? Number(data[fieldDescriptor.fieldName])
          : fieldDescriptor.defaultValue;
        break;
      case FieldType.Date:
        converted[propertyName] = data[fieldDescriptor.fieldName]
          ? new Date(data[fieldDescriptor.fieldName])
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
        const displayName: string = data[fieldDescriptor.fieldName] ? data[fieldDescriptor.fieldName] : undefined;
        if (!stringIsNullOrEmpty(displayName)) {
          if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
            // LOOKUPS --> links
            destItem.__setInternalLinks(propertyName, displayName);
            destItem[propertyName] = fieldDescriptor.defaultValue;
          }
          else {
              destItem[propertyName] = displayName;
          }
        }
        else {
            destItem[propertyName] = fieldDescriptor.defaultValue;
        }
        break;
      case FieldType.UserMulti:
        const displayNames: Array<string> = !stringIsNullOrEmpty(data[fieldDescriptor.fieldName]) ? data[fieldDescriptor.fieldName].split(";") : [];
        
        if (displayNames.length > 0) {
            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                // LOOKUPS --> links
                destItem.__setInternalLinks(propertyName, displayNames);
                destItem[propertyName] = fieldDescriptor.defaultValue;
            }
            else {
                destItem[propertyName] = displayNames;
            }
        }
        else {
            destItem[propertyName] = fieldDescriptor.defaultValue;
        }  
        break;
      case FieldType.Taxonomy:
        let termId;
        if (data[fieldDescriptor.fieldName] && data[fieldDescriptor.fieldName].includes("GP0|#")) {
          termId = data[fieldDescriptor.fieldName].split("\n").filter((str) => { return str.indexOf("GP0|#") === 0; }).map((str) => { return str.replace("GP0|#", ""); })[0];
        }
        else {
          termId = data[fieldDescriptor.fieldName]
            ? data[fieldDescriptor.fieldName].split("|")[1]
            : null;
        }
        if (!stringIsNullOrEmpty(termId)) {
          if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
              // --> links
              destItem.__setInternalLinks(propertyName, termId);
              destItem[propertyName] = fieldDescriptor.defaultValue;

          }
          else {
              destItem[propertyName] = termId;
          }

      }
      else {
          destItem[propertyName] = fieldDescriptor.defaultValue;
      }
        break;
      case FieldType.TaxonomyMulti:
        let terms;
        if (data[fieldDescriptor.fieldName] && data[fieldDescriptor.fieldName].includes("GP0|#")) {
          terms = data[fieldDescriptor.fieldName].split(";").filter((str) => { return str.indexOf("GP0|#") === 0; }).map((str) => { return str.replace("GP0|#", ""); });
        }
        else {
          terms = data[fieldDescriptor.fieldName]
            ? data[fieldDescriptor.fieldName].split(";")
            : null;
        }
        const termGuids = terms?.map(t => t.split("|")[1]).filter(t => t) || [];
        if (termGuids.length > 0) {
            if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                // LOOKUPS --> links
                destItem.__setInternalLinks(propertyName, termGuids);
                destItem[propertyName] = fieldDescriptor.defaultValue;
            }
            else {
                destItem[propertyName] = termGuids;
            }
        }
        else {
            destItem[propertyName] = fieldDescriptor.defaultValue;
        }  
        break;
      case FieldType.Json:
        converted[propertyName] = data[fieldDescriptor.fieldName]
          ? JSON.parse(data[fieldDescriptor.fieldName])
          : fieldDescriptor.defaultValue;
        break;
    }
  }

  /********************** Overrides for user field **************************************************/
  protected convertItemToDbFormat(item: T): T {
    const result: T = cloneDeep(item);
    result.cleanBeforeStorage();
    result.__clearInternalLinks();
    for (const propertyName in result) {
        if (result.hasOwnProperty(propertyName)) {
            if (this.ItemFields.hasOwnProperty(propertyName)) {
                const fieldDescriptor = this.ItemFields[propertyName];
                switch (fieldDescriptor.fieldType) {
                    case FieldType.User:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            //link defered
                            if (item[propertyName]) {
                                result.__setInternalLinks(propertyName, (item[propertyName] as unknown as User).displayName);
                            }
                            delete result[propertyName];
                        }
                        break;
                    case FieldType.Taxonomy:
                    case FieldType.Lookup:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            //link defered
                            if (item[propertyName]) {
                                result.__setInternalLinks(propertyName, (item[propertyName] as unknown as BaseItem<string | number>).id);
                            }
                            delete result[propertyName];
                        }
                        break;
                    case FieldType.UserMulti:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            const displayNames = [];
                            if (item[propertyName]) {
                                (item[propertyName] as unknown as User[]).forEach(element => {
                                    if (!stringIsNullOrEmpty(element?.displayName)) {
                                      displayNames.push(element.displayName);
                                    }
                                });
                            }
                            if (displayNames.length > 0) {
                                result.__setInternalLinks(propertyName, displayNames.length > 0 ? displayNames : []);
                            }
                            delete result[propertyName];
                        }
                        break;
                    case FieldType.TaxonomyMulti:
                    case FieldType.LookupMulti:
                        if (!stringIsNullOrEmpty(fieldDescriptor.modelName)) {
                            const ids = [];
                            if (item[propertyName]) {
                                (item[propertyName] as unknown as BaseItem<string | number>[]).forEach(element => {
                                    if (element?.id) {
                                        ids.push(element.id);
                                    }
                                });
                            }
                            if (ids.length > 0) {
                                result.__setInternalLinks(propertyName, ids.length > 0 ? ids : []);
                            }
                            delete result[propertyName];
                        }
                        break;
                    default:
                        break;
                }
            } else if (typeof (result[propertyName]) === "function") {
                delete result[propertyName];
            }
        }
    }
    return result;
}

protected async populateLinkedFields(items: T[], loadLinked?: string[], innerItems?: { [modelName: string]: BaseItem<string | number>[]; }): Promise<void> {
  await super.populateLinkedFields(items, loadLinked, innerItems);        
  // get linked fields
  const linkedFields = this.linkedFields(loadLinked).filter(lf => lf.fieldType === FieldType.User || lf.fieldType === FieldType.UserMulti);
  // init values and retrieve all ids by model
  const allDisplayNames = {};
  const innerResult = {};
  for (const key in linkedFields) {
      if (linkedFields.hasOwnProperty(key)) {
          const fieldDesc = linkedFields[key];
          allDisplayNames[fieldDesc.modelName] = allDisplayNames[fieldDesc.modelName] || [];
          const displayNames = allDisplayNames[fieldDesc.modelName];
          items.forEach((item: T) => {
              const links = item.__getInternalLinks(key);
              //init value 
              item[key] = fieldDesc.defaultValue;
              
              if (fieldDesc.fieldType === FieldType.User &&
                  // lookup has value
                  !stringIsNullOrEmpty(links)) {
                  // check in preloaded
                  let inner = undefined;
                  if (innerItems && innerItems[fieldDesc.modelName]) {
                      inner = find(innerItems[fieldDesc.modelName], ii => (ii as User).displayName === links);
                  }
                  // inner found
                  if (inner) {
                      innerResult[fieldDesc.modelName] = innerResult[fieldDesc.modelName] || [];
                      innerResult[fieldDesc.modelName].push(inner);
                  }
                  else {
                    displayNames.push(links);
                  }
              }
              else if (fieldDesc.fieldType === FieldType.UserMulti &&
                  links &&
                  links.length > 0) {
                  links.forEach((displayName) => {
                      let inner = undefined;
                      if (innerItems && innerItems[fieldDesc.modelName]) {
                          inner = find(innerItems[fieldDesc.modelName], ii => (ii as User).displayName === displayName);
                      }
                      // inner found
                      if (inner) {
                          innerResult[fieldDesc.modelName] = innerResult[fieldDesc.modelName] || [];
                          innerResult[fieldDesc.modelName].push(inner);
                      }
                      else {
                        displayNames.push(displayName);
                      }
                  });
              }
          });

      }
  }
  const resultItems: { [modelName: string]: User[] } = innerResult;
  
  // Init queries       
  const promises: Array<() => Promise<User[]>> = [];
  for (const modelName in allDisplayNames) {
      if (allDisplayNames.hasOwnProperty(modelName)) {
          const displayNames = allDisplayNames[modelName];
          if (displayNames && displayNames.length > 0) {
              const options: IBaseSPServiceOptions = {};
              // for sp services
              if(this.serviceOptions.hasOwnProperty('baseUrl')) {
                  options.baseUrl = (this.serviceOptions as IBaseSPServiceOptions).baseUrl;
              }
              const service = ServiceFactory.getServiceByModelName(modelName, options);
              promises.push(() => (service as UserService).getByDisplayNames(displayNames));
          }
      }
  }
  // execute and store
  const results = await UtilsService.executePromisesInStacks(promises, 3);
  results.forEach(itemsTab => {
      if (itemsTab.length > 0) {
          resultItems[itemsTab[0].constructor["name"]] = resultItems[itemsTab[0].constructor["name"]] || [];
          resultItems[itemsTab[0].constructor["name"]].push(...itemsTab);
      }
  });

  // Associate to items
  for (const propertyName in linkedFields) {
      if (linkedFields.hasOwnProperty(propertyName)) {
          const fieldDesc = linkedFields[propertyName];
          const refCol = resultItems[fieldDesc.modelName];
          items.forEach((item: T) => {
              const links = item.__getInternalLinks(propertyName);
              if (fieldDesc.fieldType === FieldType.User &&
                  !stringIsNullOrEmpty(links)) {
                  const litem = find(refCol, { displayName: links });
                  if (litem) {
                      item[propertyName] = litem;
                  }

              }
              else if (fieldDesc.fieldType === FieldType.UserMulti &&
                  links &&
                  links.length > 0) {
                  item[propertyName] = [];
                  links.forEach((dn) => {
                      const litem = find(refCol, { displayName: dn });
                      if (litem) {
                          item[propertyName].push(litem);
                      }
                  });
              }
          });
      }
    }
  }


  private getLogicalSequence(sequence: ILogicalSequence<T>): string {
    const cloneSequence = cloneDeep(sequence);
    if (!cloneSequence.children || cloneSequence.children.length === 0) {
      return "";
    }
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
        result = ` ${field.fieldName.toString()}:(a* OR b* OR c* OR d* OR e* OR f* OR g* OR h* OR i* OR j* OR k* OR l* OR m* OR n* OR o* OR p* OR q* OR r* OR s* OR t* OR u* OR v* OR w* OR x* OR y* OR z* OR 1* OR 2* OR 3* OR 4* OR 5* OR 6* OR 7* OR 8* OR 9* OR 0*)`;

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
