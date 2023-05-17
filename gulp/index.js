// dev deps in spfx package
const TerserPlugin = require('terser-webpack-plugin');

// From node
const fs = require('fs');
const glob = require('glob');
const path = require('path');

const TerserPluginVersion = (() => {
    try {
        const terserPath = require.resolve('terser-webpack-plugin');
        const rootTerser = path.dirname(path.dirname(terserPath));
        const packageJSON = path.join(rootTerser, 'package.json');

        let label = '';
        let [major = '0', minor = '0', fix = '0'] = require(packageJSON).version.split('.');
        [fix = '0', ...label] = fix.split('-');
        label = label?.join('-') ?? '';

        [major, minor, fix] = [major, minor, fix].map(Number);

        return { major, minor, fix, label };
    } catch (error) {
        // known version transversal used on old spfx version
        return { major: 1, minor: 4, fix: 5, label: '' };
    }
})();

// dev deps
var es = require('event-stream'), PluginError = require('plugin-error');

const stream = function (injectMethod) {
    return es.map(function (file, cb) {
        try {
            file.contents = Buffer.from(injectMethod(String(file.contents)));
            cb(null, file);
        } catch (err) {
            return cb(new PluginError('gulp-configure-dataservice', err));
        }
    });
};

const baseItems = ["BaseFile", "BaseItem", "RestFile", "RestItem", "SPFile", "SPItem", "TaxonomyTerm", "TaxonomyHidden", "User", "Entity"];

const inject = function (imports, filePath) {
    const begin = "//inject:imports", end = "//endinject";
    const comment = "\n/************************* Automatic services declaration injection for base-data-services *************************/"
    return stream(function (fileContents) {

        let importString = "";
        const declarations = [];
        for (const className in imports) {
            if (imports.hasOwnProperty(className)) {
                const classRef = imports[className];
                declarations.push(className);
                if (classRef.isFile) {
                    const dirPath = path.dirname(filePath);
                    importString += `import { ${className} } from "${path.relative(dirPath, classRef.ref).replace(/\\/g, "/")}";\n`
                }
                else {
                    importString += `import { ${className} } from "${classRef.ref}";\n`
                }

            }
        }
        const str = importString + `\nconsole.groupCollapsed("spfx-base-data-services - register services");\n[\n${declarations.map(d => "\t" + d).join(",\n")}\n].forEach(function (value) { \n\tconsole.log((value as new() => BaseService).name + " added to ServiceFactory");\n});\nconsole.groupEnd();\n`;
        const regex = new RegExp(begin + ".*" + end, 's');
        if (fileContents.match(regex)) {
            return fileContents.replace(new RegExp(begin + ".*" + end, 's'), begin + comment + "\n" + str + end);
        }
        else {
            const match = fileContents.match(/\s*import.*from.*;/g);
            if (match) {
                const matchstr = match.pop();
                const idx = fileContents.lastIndexOf(matchstr);
                return fileContents.slice(0, idx + matchstr.length) + "\n" + begin + comment + "\n" + str + end + "\n" + fileContents.slice(idx + matchstr.length);
            }
            else {
                return begin + comment + "\n" + str + end + "\n" + fileContents;
            }
        }
    });
}

function getInjectionTask(build, basePath) {
    const tsconfig = require(path.resolve(basePath, "tsconfig.json"));
    // inject dataservices in entry points to ensure decorators are applied for ServiceFactory registration (for both service and associated model)
    let injectServices = build.subTask('services-inject', function (gulp, buildOptions, done) {
        build.log("Inject services imports for ServiceFactory init");
        // find entry points in package config
        const sources = [];
        const pkgconfig = require(path.resolve(basePath, "config/config.json"));
        for (const key in pkgconfig.bundles) {
            if (pkgconfig.bundles.hasOwnProperty(key)) {
                const bundle = pkgconfig.bundles[key];
                bundle.components.forEach(component => {
                    sources.push(component.entrypoint.replace(/^\.\/lib\/(.*)\.js/, "./src/$1.ts"));
                });

            }
        }
        // construct injection
        const imports = {
            BaseService: {
                isFile: false,
                ref: "spfx-base-data-services"
            },
            TaxonomyHiddenListService: {
                isFile: false,
                ref: "spfx-base-data-services"
            },
            UserService: {
                isFile: false,
                ref: "spfx-base-data-services"
            },
            EntityService: {
                isFile: false,
                ref: "spfx-base-data-services"
            }
        };
        // search services in each tsconfig includes
        tsconfig.include.forEach((pattern) => {
            glob.sync(pattern).forEach((filePath) => {
                const buf = fs.readFileSync(filePath, "utf-8");
                const serviceDeclaration = buf.match(/@.*dataService\((["']\w+["'])?\).*export\s*class\s*(\w+)\s*extends.*/s);
                if (serviceDeclaration && serviceDeclaration.length === 3) {
                    const className = serviceDeclaration[2];
                    imports[className] = {
                        isFile: true,
                        ref: path.resolve(basePath, filePath.replace(/^src(\/.*)\.ts$/g, `src$1`))
                    }

                }
            });
        });
        es.concat(
            sources.map(source => {
                const filePath = path.resolve(basePath, source);
                const fileDir = path.dirname(filePath);
                return gulp.src(filePath).pipe(
                    inject(imports, filePath)
                ).pipe(
                    gulp.dest(fileDir)
                );

            })
        ).on('end', () => { done(); });
    });
    return injectServices;
}

function mergeWebPackConfig(build, config, basePath, includeSourceMap, sourceMapExclusions, additionnalReservedNames, afterMergeConfig) {
    const tsconfig = require(path.resolve(basePath, "tsconfig.json"));
    if (includeSourceMap) {
        // include sourcemaps for dev
        if (!build.getConfig().production) {
            build.log("Including sourcemaps in bundle");
            config.module.rules.push({
                test: /\.(js|mjs|jsx|ts|tsx)$/,
                use: ['source-map-loader'],
                exclude: sourceMapExclusions || [],
                enforce: 'pre',
            });
        }
    }
    // only prod buid
    if (build.getConfig().production) {
        build.log("Exclude services and models class names for uglify plugin");
        additionnalReservedNames = additionnalReservedNames || [];
        const reserved = ["BaseDataService", "BaseDbService", "BaseFileService", "BaseListItemService", "BaseRestService", "BaseService", "BaseTermsetService", "SPFile", "TaxonomyTerm", "TaxonomyHiddenListService", "SPFile", "TaxonomyTerm", "TaxonomyHiddenListService", "TaxonomyHidden", "UserService", "User", "Entity", "BaseItem", "RestItem", "SPItem", "RestFile", "BaseFile"].concat(additionnalReservedNames);
        const baseClasses = [];
        tsconfig.include.forEach((pattern) => {
            glob.sync(pattern).forEach((filePath) => {
                const buf = fs.readFileSync(filePath, "utf-8");
                const serviceDeclaration = buf.match(/@.*(dataService|dataModel)\((["']\w+["'])?\).*export\s*class\s*(\w+)\s*(implements\s*\w+\s*)?extends\s*(\w+)\s*.*/s);
                if (serviceDeclaration && serviceDeclaration.length === 6) {
                    const className = serviceDeclaration[3];
                    if (serviceDeclaration[1] == "dataModel") {
                        build.verbose("model class : " + className);
                        const baseClass = serviceDeclaration[5];
                        if (baseItems.indexOf(baseClass) === -1 && baseClasses.indexOf(baseClass) === -1) {
                            baseClasses.push(baseClass);
                        }
                    }
                    else {
                        build.verbose("service class : " + className);
                    }
                    reserved.push(className);
                }
            });
        });
        const addedParents = [].concat(...baseClasses);
        while (baseClasses.length > 0) {
            tsconfig.include.forEach((pattern) => {
                glob.sync(pattern).forEach((filePath) => {
                    const buf = fs.readFileSync(filePath, "utf-8");
                    const classPattern = buf.match(/.*export\s*(abstract\s*)?class\s*(\w+)\s*(implements\s*\w+\s*)?extends\s*(\w+)\s*.*/s); // /!\ implements
                    let isParentClass = false;
                    if (classPattern && classPattern.length === 5) {
                        const parentName = classPattern[2];
                        const parentBaseType = classPattern[4];
                        const nameIdx = baseClasses.indexOf(parentName)
                        isParentClass = nameIdx !== -1;
                        if (isParentClass) {
                            build.verbose("parent class : " + parentName);
                            addedParents.push(parentName);
                            baseClasses.splice(nameIdx, 1);
                            if (reserved.indexOf(parentName) === -1) {
                                reserved.push(parentName);
                            }
                            if (addedParents.indexOf(parentBaseType) === -1 && baseItems.indexOf(parentBaseType) === -1) {
                                baseClasses.push(parentBaseType);
                            }
                        }
                    }
                });
            });
        }


        build.verbose("Reserved names : " + reserved.join(", "));
        if (TerserPluginVersion.major >= 2) {
            TerserPlugin.isWebpack4 = () => true;
            function escapeRegExp(string) {
                return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
            }
            const reservedRegex = new RegExp(`^${reserved.map(escapeRegExp).join('|')}$`);
            config.optimization.minimizer = [
                new TerserPlugin({
                    extractComments: false,
                    parallel: false,
                    terserOptions: {
                        keep_classnames: reservedRegex,
                        keep_fnames: reservedRegex,
                        sourceMap: false,
                        format: { comments: false },
                        mangle: {
                            reserved,
                            keep_classnames: reservedRegex,
                            keep_fnames: reservedRegex,
                        },
                    },
                }),
            ];
        }
        else {
            config.optimization.minimizer = [
                new TerserPlugin({
                    extractComments: false,
                    sourceMap: false,
                    cache: false,
                    parallel: false,
                    terserOptions: {
                        output: { comments: false },
                        compress: { warnings: false },
                        mangle: { reserved }, // rem sample from doc ['$super', '$', 'exports', 'require']
                    },
                }),
            ];
        }
    }
    // alias
    config.resolve = config.resolve || { modules: ['node_modules'] };
    config.resolve.alias = {};
    if (tsconfig.compilerOptions.paths) {
        build.log("Include aliases in bundle");
        var baseUrl = ".";
        // get base url
        if (tsconfig.compilerOptions.baseUrl) {
            baseUrl = tsconfig.compilerOptions.baseUrl.replace(/\/*$/g, '');
        }
        // transform paths to resolve
        for (const key in tsconfig.compilerOptions.paths) {
            if (tsconfig.compilerOptions.paths.hasOwnProperty(key)) {
                if (tsconfig.compilerOptions.paths[key] && tsconfig.compilerOptions.paths[key].length > 0) {
                    var tspath = tsconfig.compilerOptions.paths[key][0];
                    var parts = tspath.split("/");
                    var lastPart = parts.pop();
                    while (lastPart === "*" || lastPart === "") {
                        lastPart = parts.pop();
                    }
                    tspath = parts.join("/") + "/" + lastPart;
                    var destpath = baseUrl + "/" + tspath.replace(/^\/*|\/*$/g, '');
                    // folder
                    if (lastPart.indexOf(".") === -1) {
                        destpath += "/";
                    }
                    else {
                        destpath = destpath.replace(/\.ts$/g, '.js');
                    }
                    destpath = destpath.replace("/src/", "/lib/");
                    // Remove /* from key if needed
                    var destkey = key.replace(/^(.*)\/\*$/g, "$1");
                    config.resolve.alias[destkey] = path.resolve(basePath, destpath);
                }
            }
        }
        for (const key in config.resolve.alias) {
            if (config.resolve.alias.hasOwnProperty(key)) {
                build.verbose("Alias " + key + " --> " + config.resolve.alias[key]);
            }
        }
    }
    if (afterMergeConfig) {
        build.log("Running addintionnal config");
        afterMergeConfig(config);
    }
    return config;
}

function configureSpfxProject(build, basePath, includeSourceMap, sourceMapExclusions, additionnalReservedNames, afterMergeConfig) {
    const injectServices = getInjectionTask(build, basePath);
    build.rig.addPreBuildTask(injectServices);

    /*
    * modify webpack config :
    *    - to avoid uglyfying reserved names (services & models)
    *    - to handle aliases
    *    - to link sourcmaps
    */
    build.configureWebpack.mergeConfig({
        additionalConfiguration: (config) => {
            return mergeWebPackConfig(build, config, basePath, includeSourceMap, sourceMapExclusions, additionnalReservedNames, afterMergeConfig);
        }
    });
}

module.exports = {
    getInjectionTask: getInjectionTask,
    mergeWebPackConfig: mergeWebPackConfig,
    configureSpfxProject: configureSpfxProject
};
