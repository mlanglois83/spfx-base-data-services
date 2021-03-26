export class LoggingService {
    public static addLoggingToTaggedMembers(instance: any, logFormat?: string): void {  
        const instanceConstructor = instance.constructor;
        const className = instanceConstructor["name"] || "object"; 
        if(instanceConstructor.tracedMembers && instanceConstructor.tracedMembers.length > 0) {
            instanceConstructor.tracedMembers.forEach(tracedMember => {
                instance[tracedMember] = LoggingService.getLoggableFunction(instance, className, instance[tracedMember], tracedMember, logFormat);
            });
        }        
    }

    public static addLoggingToFullInstance(instance: any, logFormat?: string): void {  
        const instanceConstructor = instance.constructor;
        const className = instanceConstructor["name"] || "object"; 
        for(const name in instance){
            if(name !== "constructor") {
                const potentialFunction = instance[name];            
                if(Object.prototype.toString.call(potentialFunction) === '[object Function]'){
                    instance[name] = LoggingService.getLoggableFunction(instance, className, potentialFunction, name, logFormat);                    
                }
            }            
        }
    }

    
    private static getLoggableFunction(instance: any, className, func: any, name: string, logFormat?: string): any {
        // override function for logging with duration
        return  function(...args: any[]): any {            
            const startDate = new Date();
            const result = func.apply(instance, args);
            // case of async function
            if(result instanceof Promise) {
                result.then((): any => {
                    LoggingService.log(instance, className, startDate, name, args, logFormat);
                });
            }
            // classical function
            else {
                LoggingService.log(instance, className, startDate, name, args);
            }
            return result;
        };
    } 

    public static readonly defaultLogFormat = "%Time% - [%ClassName%] --> %Function%: %Duration%ms";

    private static log(instance: any, className: string, startDate: Date, functionName: string, args: any[], logFormat?: string): void {
        logFormat = logFormat || LoggingService.defaultLogFormat;

        const endDate = new Date();
        const duration = new Date().getTime() - startDate.getTime();
        const functionString = `${functionName}(${args ? args.map(arg => LoggingService.formatValueForFunction(arg)).join(", ") : ""})`;

        let logText = logFormat.replace(/%Time%/g, new Date().toLocaleTimeString());
        logText = logText.replace(/%Date%/g, new Date().toLocaleDateString());
        logText = logText.replace(/%DateTime%/g, new Date().toLocaleString());        
        logText = logText.replace(/%Start%/g, startDate.toLocaleString());        
        logText = logText.replace(/%End%/g, endDate.toLocaleString());      
        logText = logText.replace(/%ClassName%/g, className);      
        logText = logText.replace(/%Function%/g, functionString);      
        logText = logText.replace(/%Duration%/g, duration.toString());
        // additionnal props
        logText = logText.replace(/%Property:[^%]+%/g, (match) => {
            let result = "";
            const logprop = match.replace(/%Property:(.*)%/g, "$1");
            const parts = logprop.split(".");
            if(parts.length > 0) {
                let val: any = instance;
                parts.forEach(p => {
                    if(val !== undefined && val !== null) {
                        val = val[p];
                    }
                });
                result = LoggingService.formatValue(val);
            }
            return result;
        });
        console.log(logText);
    }

    private static formatValue(value: any): string {
        if(value) {
            if(typeof(value) === "object") {
                return JSON.stringify(value);
            }
        }
        return `${value}`;
    }

    private static formatValueForFunction(value: any, includeCount = true): string {
        if(value) {
            if(typeof(value) === "object") {
                if(Array.isArray(value)) {
                    if(value.length > 0) {
                        // format first object
                        let itemFormated: string = typeof(value[0]);
                        if(itemFormated === "object") {
                            itemFormated = LoggingService.formatValueForFunction(value[0], false);
                        }
                        return `Array<${itemFormated}>` + (includeCount ? `[${value.length}]` : "");
                    }
                    else {
                        return "[]";
                    }
                }
                else {
                    // return object type
                    return value?.constructor?.name || "object";

                }
            }
        }
        return `${value}`;
    }

}