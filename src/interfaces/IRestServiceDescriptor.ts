import { IEndPointBindings } from ".";

export interface IRestServiceDescriptor {
    relativeUrl: string;
    endpoints: IEndPointBindings;
}