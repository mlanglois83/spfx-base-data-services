
[![Total alerts](https://img.shields.io/lgtm/alerts/g/mlanglois83/spfx-base-data-services.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/mlanglois83/spfx-base-data-services/alerts/)
[![Language grade: JavaScript](https://img.shields.io/lgtm/grade/javascript/g/mlanglois83/spfx-base-data-services.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/mlanglois83/spfx-base-data-services/context:javascript)

# spfx-base-data-services

## Contents

- [spfx-base-data-services](#spfx-base-data-services)
  - [Contents](#contents)
  - [Description](#description)
  - [Installation](#installation)
  - [Project integration](#project-integration)
    - [Initialization](#initialization)
    - [Packaging the solution](#packaging-the-solution)
  - [Implementation](#implementation)
    - [Service Factory](#service-factory)
    - [SharePoint List](#sharepoint-list)
      - [List item model](#list-item-model)
      - [List service](#list-service)
    - [Taxonomy Term set](#taxonomy-term-set)
      - [Taxonomy term model](#taxonomy-term-model)
      - [Term set service](#term-set-service)
    - [Library](#library)
      - [Library service](#library-service)
      - [File model](#file-model)
    - [Extending Services](#extending-services)
    - [Overriding default services methods](#overriding-default-services-methods)
    - [Creating a custom service](#creating-a-custom-service)
    - [Synchronization events](#synchronization-events)
  - [Using a service](#using-a-service)
  - [Classes and interfaces description](#classes-and-interfaces-description)

## Description

spfx-base-dataservice  is a set of base classes and tools aimed to create a data service able to interract with main SharePoint / O365 data sources in SPFX web parts. It contains base implementation for the following services:

- SharePoint List including main data types (text, boolean, lookup, Taxonomy...).
- Document Library.
- SharePoint Users.
- Taxonomy Term set stored in a global group or in the site collection group.

The main feature provided by implementing this library is:

- Easy models and service design.
- Automatic link management for linked fields (Taxonomy, Lookups, Users).
- Cache management: all service can store data in cache using an age in minutes.
- Online/Offline management for Progressive Web Apps: calls to service automatically switch to local storage in case the app is offline. All features are available.
- Version conflicts check (version management must be activated on list). In case the destination item is newer, an error is thrown and the destination item is returned.
- Automatic data synchronization. All actions made by services offline are stored in a transactions table. A synchronization process runs api calls for all transactions and guarantees data consistency between SharePoint and local database.

## Installation

To use the package, it must be installed via npm and saved as project dependency using the command:

`npm i spfx-base-data-service --save`

## Project integration

  The service needs a few configuration to work in a SPFX webpart project.
  
### Initialization
  
  In webparts code file override init method as following :
  
```javascript
public  onInit(): Promise<void> {
    return  super.onInit().then(_  => {
        ServicesConfiguration.Init({
            lastConnectionCheckResult:  false,
            dbName:  "sp-db",
            dbVersion:  1,
            checkOnline:  true, // true for pwa, else false
            context:  this.context,
            currentUserId: -1, // if used in queries, must be initialized after
            serviceFactory:  new  ServiceFactory(), // service factory implementation
            tableNames:  ["list", "termset"], // 1 table per service
            translations: { // translations used for synchronization
                AddLabel:  strings.AddLabel,
                DeleteLabel:  strings.DeleteLabel,
                IndexedDBNotDefined:  strings.IndexedDBNotDefined,
                SynchronisationErrorFormat:  strings.SynchronisationErrorFormat,
                UpdateLabel:  strings.UpdateLabel,
                UploadLabel:  strings.UploadLabel,
                versionHigherErrorMessage:  strings.versionHigherErrorMessage,
                typeTranslations: { // One per model, key is class name
                    SPFile:  strings.SPFileLabel
                }
            }
        });
    });
}
```
  
### Packaging the solution
  
  As services use Class names to generate model and service instances, the following function must be added in file gulpfile.js:
  
``` javascript
...
const  TerserPlugin = require('terser-webpack-plugin');
const  glob = require('glob');

build.configureWebpack.setConfig({
    additionalConfiguration: (config) => {
        // only prod buid
        if (build.getConfig().production) {
            // get excluded names for uglify
            const reserved = glob.sync('./src/{services,models}/**/*.ts').map((filePath) => {
                return filePath.replace(/.*\/(\w+)\.ts/g, "$1");
            }).concat("SPFile", "TaxonomyTerm", "TaxonomyHiddenListService", "TaxonomyHidden", "UserService", "User");
            config.optimization.minimizer =
            [
                new TerserPlugin
                (
                    {
                        extractComments: false,
                        sourceMap: false,
                        cache: false,
                        parallel: false,
                        terserOptions:
                        {
                            output: { comments: false },
                            compress: { warnings: false },
                            mangle: {
                                reserved: reserved
                            }
                        }
                    }
                )
            ];
        }
        return config;
    }
});

```

where  `./src/{services,models}/**/*.ts` is the path in solution where services and models are stored

## Implementation

### Service Factory

To be able to instanciate models ans services, a service factory based on type [BaseServiceFactory](#baseservicefactory). The service factory exposes 2 methods :

- create for service instanciation
- getItemTypeByName for model instanciation

Each method is based on model type name, implementation of service factory should be as following:

```javascript
// import model types and service types
import { BaseServiceFactory, BaseDataService, IBaseItem, TaxonomyHidden, User, TaxonomyHiddenListService, UserService} from "spfx-base-data-services";
export class ServiceFactory extends BaseServiceFactory {
    /**
    * Instanciate a data service given associated model name
    * @param  modelName - Model associated with the service type to instanciate
    */
    public create(modelName: string): BaseDataService<IBaseItem> {
        let result = null;
        switch (modelName) {
            /* Base */
            case TaxonomyHidden["name"]:
                result = new TaxonomyHiddenListService();
                break;
            case User["name"]:
                result = new UserService();
                break;
            /*
               Add cases for all model names
            */
            default:
                throw Error(strings.errorUnknownItemTypeName);
        }
        return result;
    }

    /**
    * Retrieve model constructor given its name
    * @param modelName - Model type name
    */
    public getItemTypeByName(modelName: string): (new (item?: any) => IBaseItem) {
        let result = super.getItemTypeByName(typeName);
        if(!result) {
            switch (typeName) {
                /* Base */
                case TaxonomyHidden["name"]:
                    result = TaxonomyHidden;
                    break;
                case User["name"]:
                    result = User;
                    break;
                /*
                    Add cases for all model names
                */
                default:
                    throw Error(strings.errorUnknownItemTypeName);
            }
        }
        return result;
    }
}
```

### SharePoint List

#### List item model

A list item model is a class that inherits from base class [SPItem](#spitem). To link properties to list fields, use decorator function [spField](#spfield) with appropriate parameters according to field name, type and the way you want to retrieve value. By default, ID and Title field are retrieved in id and title properties.
In case of linked fields in model declaration, associated models and services must exist and must be declared in ServiceFactory implementation.

Sample:

```javascript
import { SPItem, FieldType, Decorators } from "spfx-base-data-services";
// import lookup types
const spField = Decorators.spField;

export class Model extends SPItem {
    //////////////////
    // Simple types //
    //////////////////
    // Text, boolean --> no type needed
    @spField({fieldName: "TextField", defaultValue: ""})
    public text: string;
    @spField({fieldName: "BooleanField", defaultValue: false})
    public bool: boolean;
    // Other simple types
    @spField({fieldName: "DateField", fieldType:FieldType.Date, defaultValue: null})
    public date: Date;
    @spField({fieldName: "JsonField", fieldType:FieldType.Json, defaultValue: null})
    public json: any;

    //////////////
    // Taxonomy //
    //////////////
    @spField({fieldName: "TaxonomyField", fieldType:FieldType.Taxonomy, modelName: "TaxoModelName", defaultValue: null})
    public taxonomyField: TaxoModelName;
    @spField({fieldName: "TaxonomyFieldMulti", fieldType:FieldType.TaxonomyMulti, modelName: "TaxoModelName", defaultValue: []})
    public taxonomyField: Array<TaxoModelName>;
    // On Taxonomy fields, if model is omitted, retrieve wssid only
    @spField({fieldName: "TaxonomyField", fieldType:FieldType.Taxonomy, defaultValue: -1})
    public taxonomyFieldWssId: number;

    ////////////
    // Lookup //
    ////////////
    @spField({fieldName: "LookupField", fieldType:FieldType.Lookup, modelName: "LookupModelName", defaultValue: null})
    public lookup: LookupModelName;
    @spField({fieldName: "LookupFieldMulti", fieldType:FieldType.LookupMulti, modelName: "LookupModelName", defaultValue: []})
    public lookupMulti: Array<LookupModelName>;
    // On Lookup fields, if model is omitted, retrieve id only
    @spField({fieldName: "LookupField", fieldType:FieldType.Lookup, defaultValue: -1})
    public lookupId: number;

    //////////
    // User //
    //////////
    @spField({fieldName: "UserField", fieldType:FieldType.User, modelName: "User", defaultValue: null})
    public user: User;
    @spField({fieldName: "UserFieldMulti", fieldType:FieldType.UserMulti, modelName: "User", defaultValue: []})
    public userMulti: Array<User>;
    // On User fields, if model is omitted, retrieve id only
    @spField({fieldName: "UserField", fieldType:FieldType.User, defaultValue: -1})
    public userId: number;
}
```

#### List service

A SharePoint List service inherits from base class BaseListItemService. Links to SharePoint list and local db are set by overriding constructor. There is no other method to declare if the solution only needs to access list via already available operations:

- Add or update
- Delete
- Get all elements
- Get by query
- Get by id(s)

Sample:

```javascript
// Add import to Model
import { BaseListItemService } from "spfx-base-data-services";

export class ListService extends BaseListItemService<Model> {
    constructor() {
        // find items in list with web relative url /lists/listurl, store cache in table ListTableName and cache results for 10 minutes
        super(Model, "/lists/listurl", "ListTableName", 10);
    }
}
```

### Taxonomy Term set

#### Taxonomy term model

A taxonomy term model is a class that inherits from base class [TaxonomyTerm](#taxonomyterm).
By default, this object exposes the following properties :

- Term id
- Label
- Term path
- Custom properties
- Custom sort order
- Deprecated

Sample:

```javascript
import { TaxonomyTerm } from  'spfx-base-data-services';

export  class  ModelName extends  TaxonomyTerm {
}
```

#### Term set service

A taxonomy term set service inherits from base class [BaseTermsetService](#basetermsetservice). Links to SharePoint term set and local db are set by overriding constructor. Service can search global term set or for local term set (stored in site collection group), link can be made on term set id or by name. There is no other method to declare if the solution only needs to access term set via already available operations:

- Get all elements
- Get by id(s)

Sample:

```javascript

import { BaseTermsetService } from 'spfx-base-data-services';
// import ModelName

export class TermSetService extends BaseTermsetService<ModelName> {
    constructor() {
        super(NameOrId, TableName, false /* store in site collection group */, 1440 /* cache duration in minutes */);
    }
}
```

### Library

#### Library service

#### File model

### Extending Services

### Overriding default services methods

### Creating a custom service

### Synchronization events

## Using a service

## Classes and interfaces description
