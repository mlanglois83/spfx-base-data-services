import { BaseService } from "../base/BaseService";
export declare class SynchronizationService extends BaseService {
    private transactionService;
    constructor();
    run(): Promise<Array<string>>;
    private formatError;
}
