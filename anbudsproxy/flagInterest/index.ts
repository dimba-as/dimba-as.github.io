
import { AzureFunction, Context, HttpRequest } from "@azure/functions";
const SPFetchClient = require("@pnp/nodejs-commonjs").SPFetchClient;
const connectSPO = require("../shared/connectSPO");
const getToken = require('../Shared/getToken');
const getUserFromGraph = require('../Shared/getUserFromGraph');
import { sp } from "@pnp/sp-commonjs";
const flagInterest: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> { 
    const webUrl = req.body.webUrl;
    const listId = req.body.listId;
    const listItemId = req.body.listItemId;
    const from = req.body.from;
    const displayName = req.body.displayName;
    await connectSPO(webUrl);
    const token = await getToken(context);
    const graphUser = await getUserFromGraph(context, token, from);
    const user = await sp.web.ensureUser(from);
    const results = await sp.web.lists.getByTitle("Foresporsler").items.add({Title:displayName,avAnbudId:listItemId, avPersonId:user.data.Id});

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
