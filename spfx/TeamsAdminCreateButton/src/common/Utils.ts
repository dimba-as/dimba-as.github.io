import { HttpClient, IHttpClientOptions, HttpClientResponse, AadHttpClient } from '@microsoft/sp-http';
//import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { ExtensionContext } from "@microsoft/sp-extension-base";

export default class Utils {

    public context: ExtensionContext;

    constructor(context: ExtensionContext) {
        this.context = context;
    }

    public async createTeam(functionUrl): Promise<string> {

        return new Promise<string>((resolve, reject) => {
            const requestHeaders: Headers = new Headers();
            requestHeaders.append('Content-Type', 'application/json');
            const postOptions: IHttpClientOptions = {
                headers: requestHeaders
               
            };

            this.context.httpClient.get(functionUrl, HttpClient.configurations.v1, postOptions)
                .then(async (res) => {
                    console.log(res);
                    let response = await res.json();
                    console.log(response);
                    resolve(response.message);
                })
                .catch(error => {
                    console.error(error);
                    reject(error);
                });
        });
    }
}