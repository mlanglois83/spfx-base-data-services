// dev deps in spfx package
const build = require('@microsoft/sp-build-web');
const TerserPlugin = require('terser-webpack-plugin');

// From node 
const fs = require('fs');
const glob = require('glob');
const path = require('path');

// dev deps
var es = require('event-stream'), PluginError = require('plugin-error');

const stream = function(injectMethod){
    return es.map(function (file, cb) {
        try {
            file.contents = Buffer.from( injectMethod( String(file.contents) ));
        } catch (err) {
            return cb(new PluginError('gulp-configure-dataservice', err));
        }
        cb(null, file);
    });
};
const inject = function(imports, filePath) {
    const begin = "//inject:imports", end = "//endinject";
    return stream(function(fileContents) {

        let importString = "";
        const declarations = [];
        for (const className in imports) {
            if (imports.hasOwnProperty(className)) {
                const classRef = imports[className];
                declarations.push(className);
                if(classRef.isFile) {
                    const dirPath = path.dirname(filePath);
                    importString += `import { ${className} } from "${path.relative(dirPath, classRef.ref).replace(/\\/g, "/")}";\n`
                }
                else {
                    importString += `import { ${className} } from "${classRef.ref}";\n`
                }
                
            }
        }
        const str = importString + `\nconsole.groupCollapsed("spfx-base-data-services - register services");\n[\n${ declarations.map(d => "\t" + d).join(",\n")}\n].forEach(function (value) { \n\tconsole.log(value["name"] + " added to ServiceFactory");\n});\nconsole.groupEnd("spfx-base-data-services - register services");\n`;
        const regex = new RegExp(begin + ".*" + end, 's');
        if(fileContents.match(regex)) {
            return fileContents.replace(new RegExp(begin + ".*" + end, 's'), begin + "\n" + str + end);
        }
        else {
            const match = fileContents.match(/\s*import.*from.*;/g);
            if(match) {                
                const matchstr = match.pop();
                const idx=fileContents.lastIndexOf(matchstr);
                return fileContents.slice(0, idx + matchstr.length) + "\n" + begin + "\n" + str + end + "\n" + fileContents.slice(idx + matchstr.length);
            }
            else {
                return begin + "\n" + str + end + "\n" + fileContents;
            }
        }
    });
}



function setConfig(basePath, includeSourceMap, afterSetConfig) { 
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
                    sources.push(component.entrypoint);
                });
                
            }
        }
        // construct injection
        const imports = {
            TaxonomyHiddenListService: { 
                isFile: false,
                ref: "spfx-base-data-services"
            },
            UserService: { 
                isFile: false,
                ref: "spfx-base-data-services"
            }
        };        
        // search services in each tsconfig includes
        tsconfig.include.forEach((pattern) => {
            glob.sync(pattern).forEach((filePath) => {
                const buf = fs.readFileSync(filePath, "utf-8");
                const serviceDeclaration = buf.match(/@.*dataService\(("\w+")?\).*export\s*class\s*(\w+)\s*extends.*/s);
                if(serviceDeclaration && serviceDeclaration.length === 3) {
                    const className = serviceDeclaration[2];
                    imports[className] = {
                        isFile: true,
                        ref: path.resolve(basePath, filePath.replace(/^src(\/.*)\.ts$/g,"lib$1.js"))
                    }
                }
            });
        });

        // inject in js file
        sources.forEach(source => {
            const filePath = path.resolve(basePath,source);
            const fileDir = path.dirname(filePath);
            gulp.src(filePath).pipe(
                inject(imports, filePath)
            ).pipe(
                gulp.dest(fileDir)
            );
            
        });
        done();
    });
    build.rig.addPostBuildTask(injectServices);

    /*
    * modify webpack config :
    *    - to avoid uglyfying reserved names (services & models) 
    *    - to handle aliases
    *    - to link sourcmaps
    */
    build.configureWebpack.setConfig({
        additionalConfiguration: (config) => {
    
            if(includeSourceMap) {                
                // include sourcemaps for dev
                if (!build.getConfig().production) {
                    build.log("Including sourcemaps in bundle");
                    config.module.rules.push({
                        test: /\.(js|mjs|jsx|ts|tsx)$/,
                        use: ['source-map-loader'],
                        enforce: 'pre',
                    });
                }
            }
            // only prod buid
            if (build.getConfig().production) {                
                build.log("Exclude services and models class names for uglify plugin");
                const reserved = ["SPFile", "TaxonomyTerm", "TaxonomyHiddenListService", "TaxonomyHidden", "UserService", "User", "BaseItem", "RestItem", "SPItem", "RestFile", "BaseFile"];
                tsconfig.include.forEach((pattern) => {
                    glob.sync(pattern).forEach((filePath) => {
                        const buf = fs.readFileSync(filePath, "utf-8");
                        const serviceDeclaration = buf.match(/@.*(dataService|dataModel)\(("\w+")?\).*export\s*class\s*(\w+)\s*extends.*/s);
                        if(serviceDeclaration && serviceDeclaration.length === 4) {
                            const className = serviceDeclaration[3];

                            reserved.push(className);
                        }
                    });
                });
                build.verbose("Reserved names : " + reserved.join(", "));
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
                                            reserved: reserved // rem sample from doc ['$super', '$', 'exports', 'require']
                                        }
                                    }
                                }
                            )
                    ];
            }
    
            // alias
            config.resolve = { alias: {}, modules: ['node_modules'] };
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
            if(afterSetConfig) {
                build.log("Running addintionnal config");
                afterSetConfig(config);
            }
            return config;
        }
    });
}
module.exports = setConfig;
