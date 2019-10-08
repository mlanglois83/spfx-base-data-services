export interface IBaseItem {
    id: number | string;
    title: string;
    queries?: Array<number>;
    version?: number;
    convert?: () => any;
    onAddCompleted?: (addResult: any) => void;
    onUpdateCompleted?: (updateResult: any) => void;
}