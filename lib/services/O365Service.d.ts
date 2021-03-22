import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISite } from "../models";
export declare class O365Service {
    context: WebPartContext;
    setup(context: WebPartContext): void;
    getGroupBySiteName(siteName: string): Promise<ISite[]>;
}
declare const GroupService: O365Service;
export default GroupService;
//# sourceMappingURL=O365Service.d.ts.map