
# spfx-base-data-services

## Contents

- [Description](#description)
- [Installation](#installation)
- [Project integration](#project-integration)
  - [Initialization](#initialization)
  - [Packaging the solution](#packaging-the-solution)

## Description

spfx-base-dataservice  is a set of base classes and tools aimed to create a data service able to interract with main SharePoint / O365 data sources in SPFX webparts. It contains base implementation for the following services:

- SharePoint List including main data types (text, boolen, lookup, Taxonomy...).
- Document Library.
- SharePoint Users.
- Taxonomy Termset stored in a global group or in the site collection group.

The main feature provided by implementing this library is:

- Easy models and service design.
- Automatic link management for linked fields (Taxonomy, Lookups, Users).
- Cache management: all service can store data in cache using an age in minutes.
- Online/Offline management for Progressive Web Apps: calls to service automatically switch to local storage in case the app is offline. All features are available.
- Version conflicts check (version management must be activated on list). In case the destination item is newer, an error is thrown and the destination item is returned.
- Automatic data synchronization. All actions made by services offline are stored in a transactions table. A synchronization process runs api calls for all transactions and garantees data consistency between SharePoint and local database.

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
            translations: { // translations used for synchronization, optional
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

### SharePoint List

#### Model

#### Service

### Taxonomy Termset

### Overriding default services methods

### Creating a custom service

### Library

## Using a service

## Classes description
