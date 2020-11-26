import { IEndPointBindings } from ".";

export interface IRestServiceDescriptor {
    relativeUrl: string;
    disableVersionCheck?: boolean;
    endpoints: IEndPointBindings;
}