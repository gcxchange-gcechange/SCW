import { MSGraphClient, SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, HttpClientResponse, HttpClient, IHttpClientOptions } from "@microsoft/sp-http";

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISite, ISiteCollection} from "../models";
import { GraphRequest } from "@microsoft/microsoft-graph-client";
import { Site } from "@microsoft/microsoft-graph-types";


import { MsalClient } from "@pnp/msaljsclient";
import { graph } from "@pnp/graph/presets/all";
import { resultContent } from "office-ui-fabric-react/lib/components/FloatingPicker/PeoplePicker/PeoplePicker.scss";

export class O365Service {
    public context: WebPartContext;

    public setup(context: WebPartContext): void {
      this.context = context;
    }

    public getGroupBySiteName(siteName: string): Promise<ISite[]>{
        return new Promise<ISite[]>((resolve, reject) => {
          try {
            // Prepare the output array
            var sites: Array<ISite> = new Array<ISite>();
    
            this.context.msGraphClientFactory
              .getClient()
              .then((client: MSGraphClient) => {
                client
                  .api(`/sites?search='${siteName}'`)
                  .get((error: any, groups: ISiteCollection, rawResponse: any) => {
                    // Map the response to the output array
                    groups.value.map((item: any) => {
                        sites.push({
                        id: item.id,
               
                      });
                    });
    
                    
                    console.log(sites);
                    resolve(sites);
                  });
              });
          } catch (error) {
            console.error(error);
          }
        });
      }


}

const GroupService = new O365Service();
export default GroupService;