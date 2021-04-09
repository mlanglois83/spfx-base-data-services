// dev deps in spfx package
const build = require('@microsoft/sp-build-web');
const TerserPlugin = require('terser-webpack-plugin');

// From node 
const fs = require('fs');
const glob = require('glob');
const path = require('path');

// dev deps
const es = require('event-stream')
const PluginError = require('plugin-error');

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
const replace = function(begin, end, str) {
    return stream(function(fileContents) {
      return fileContents.replace(new RegExp(begin + ".*" + end, 's'), begin + str + end);
    });
}



function setConfig(basePath, includeSourceMap, afterSetConfig) {
    const tsconfig = require(path.resolve(basePath, "tsconfig.json"));
    let injectServices = build.subTask('services-inject', function (gulp, buildOptions, done) {        
        build.log("Inject services imports for ServiceFactory init");
        const target = gulp.src(path.resolve(__dirname,"../dist/index.js"));
        // construct injection
        let imports = `\nimport { TaxonomyHiddenListService, UserService } from "./services";\n`;
        let classNames = "\nTaxonomyHiddenListService;\nUserService;\n";
        tsconfig.include.forEach((pattern) => {
                glob.sync(pattern).forEach((filePath) => {
                    const buf = fs.readFileSync(filePath, "utf-8");
                    const serviceDeclaration = buf.match(/@.*dataService\(("\w+")?\).*export\s*class\s*(\w+)\s*extends.*/s);
                    if(serviceDeclaration && serviceDeclaration.length === 3) {
                        const className = serviceDeclaration[2];
                        imports += `import { ${className} } from "../../../${filePath.replace(/^src(\/.*)\.ts$/g,"lib$1.js")}";\n`;
                        classNames += `${className};\n`;
                    }
                });
            });
        
        target.pipe(
            replace("//inject:imports", "//endinject", imports + classNames)
        ).pipe(
            gulp.dest(path.resolve(__dirname,"../dist"))
        );
        done();
    });
    build.rig.addPostBuildTask(injectServices);
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
