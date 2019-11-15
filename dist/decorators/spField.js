export function spField(declaration) {
    return function (target, propertyKey) {
        if (!target.constructor.Fields) {
            target.constructor.Fields = {};
        }
        target.constructor.Fields[propertyKey] = declaration;
    };
}
//# sourceMappingURL=spField.js.map