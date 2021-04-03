export enum TraceLevel {
    None = 0,
    Service = 1,
    Internal = 2,
    ServiceUtilities = 4,
    Queries = 8,
    DataBase = 16,
    Custom = 32,
    ALL = 63
}