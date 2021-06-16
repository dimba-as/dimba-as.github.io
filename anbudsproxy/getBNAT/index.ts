
import { AzureFunction, Context, HttpRequest } from "@azure/functions";
const fetch = require('node-fetch');
const getBNAT: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const url = 'https://login.bisnode.com/as/token.oauth2';
    const cs = 'a408aeab-314f-4080-b861-b5a966f2d43c:313b33b4-3bb5-4503-8902-2c5356fd5b3f';
    const auth = 'Basic ' + Buffer.from(cs).toString('base64');


    let myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/x-www-form-urlencoded");
    myHeaders.append("Authorization", auth);

    const raw = "grant_type=client_credentials&scope=credit_data_companies";

    const requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: raw,
        redirect: 'follow'
    };

    const result = await fetch("https://login.bisnode.com/as/token.oauth2", requestOptions);
    const data = await result.text()
      
    context.res = {
        // status: 200, /* Defaults to 200 */ 
        body: JSON.parse(data).access_token
    };
      
};

export default getBNAT;
