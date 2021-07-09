
import { AzureFunction, Context, HttpRequest } from "@azure/functions";
const SPFetchClient = require("@pnp/nodejs-commonjs").SPFetchClient;
const connectSPO = require("../shared/connectSPO");
const getToken = require('../Shared/getToken');
const getUserFromGraph = require('../Shared/getUserFromGraph');
import { sp } from "@pnp/sp-commonjs";
import { Web } from "sp-pnp-js";
const flagInterest: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> { 
    const webUrl = req.body.webUrl;
    const listItemId = req.body.listItemId;
    const from = req.body.from;
    const displayName = req.body.displayName;
    context.log(webUrl);
    await connectSPO(webUrl);
    const w = await sp.web.get();

    context.log(JSON.stringify(w, null, 4));
    context.log(w.Title);
    const token = await getToken(context);
    const graphUsers = await getUserFromGraph(context, token, from);
    let graphUserCompany = "";
    let graphUserPrincipalName = "";
    if(graphUsers && graphUsers.value && graphUsers.value.length===1 ){
        graphUserCompany = graphUsers.value[0].companyName;
        graphUserPrincipalName = graphUsers.value[0].userPrincipalName;
    }
    
    const list = await sp.web.lists.getByTitle("Forespørsler");
    const results = await list.items.add({Title:displayName,avSelskap:graphUserCompany,avAnbudId:listItemId, avPersonEmail:from, avUserPrincipalName:graphUserPrincipalName});

    
    if (results) {
        context.res = {
            // status: 200, /* Defaults to 200 */ 
            body: JSON.stringify(results)
        };
    }
    else {
        context.res = {
            status: 400,
            body: "No contacts found"
        };
    }
};

export default flagInterest;
