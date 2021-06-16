
import { AzureFunction, Context, HttpRequest } from "@azure/functions";
const fetch = require('node-fetch');
const getBNAT: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const regNumber = context.req.query.regnumber;
    const url = 'https://login.bisnode.com/as/token.oauth2';
    const cs = 'a408aeab-314f-4080-b861-b5a966f2d43c:313b33b4-3bb5-4503-8902-2c5356fd5b3f';
    const auth = 'Basic ' + Buffer.from(cs).toString('base64');


    let atHeaders = new Headers();
    atHeaders.append("Content-Type", "application/x-www-form-urlencoded");
    atHeaders.append("Authorization", auth);

    const atRaw = "grant_type=client_credentials&scope=credit_data_companies";

    const atRequestOptions = {
        method: 'POST',
        headers: atHeaders,
        body: atRaw,
        redirect: 'follow'
    };

    const result = await fetch(url, atRequestOptions);
    const data = await result.text();
    const accessToken = JSON.parse(data).access_token;

    //GET CR
    var crHeaders = new Headers();
    crHeaders.append("Authorization", "Bearer " + accessToken);
    crHeaders.append("Content-Type", "application/json");

    var raw = JSON.stringify({
        "registrationNumber": regNumber,
        "reference": "string",
        "onBehalfOf": "string",
        "language": "NO",
        "segments": [
            "RISK"
        ]
    });

    var requestOptions = {
        method: 'POST',
        headers: crHeaders,
        body: raw,
        redirect: 'follow'
    };

    const crResult = await fetch("https://sandbox-api.bisnode.com/credit-data-companies/v2/companies/no", requestOptions);
    const crData = await crResult.text();

    context.res = {
        // status: 200, /* Defaults to 200 */ 
        body: crData
    };

};

export default getBNAT;
