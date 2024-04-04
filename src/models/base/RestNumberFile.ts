import { RestFile } from "./RestFile";

export abstract class RestNumberFile extends RestFile<number> {
    public get typedKey(): number {
        return 0;
    }
}