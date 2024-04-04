import { RestFile } from "./RestFile";

export abstract class RestStringFile extends RestFile<string> {
    public get typedKey(): string {
        return "";
    }
}