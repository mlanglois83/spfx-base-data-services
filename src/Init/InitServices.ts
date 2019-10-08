import { IConfiguration } from "../interfaces";
import { BaseService } from "../services/base/BaseService";

export function initServices(config: IConfiguration) {
    BaseService.Init(config);
}