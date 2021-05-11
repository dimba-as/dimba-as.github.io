
import { AzureFunction, Context, HttpRequest } from "@azure/functions";
const SPFetchClient = require("@pnp/nodejs-commonjs").SPFetchClient;
const connectSPO = require("../shared/connectSPO");
import { sp } from "@pnp/sp-commonjs";


import { ISearchQuery, SearchResults, SearchQueryBuilder } from "@pnp/sp/search";
const getOpenListings: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    await connectSPO("https://dimbaas.sharepoint.com/");
    
    const results2 = await sp.search(<ISearchQuery>{
        Querytext: "ContentTypeId:0x0100EA4AF1BB52A44A48B0A8D8EBDE86B2F900826E760165E09246A64BC70214D36943 avPublisert=1 avPubliseringsniva=Offentlig",
        RowLimit: 10,
        SelectProperties:[
            "Title",
           "avCustomFag",
           "avCustomAnbudsbeskrivelse",
           "avPublisert",
           "avPubliseringsniva",
           "avAnbudsfristOWSDATE",
           "avOppstartAarOWSNMBR",
           "avOppstartsperiodeOWSCHCS",
           "avFerdigstillelseperiodeOWSCHCS",
           "avFerdigstillelseAarOWSNMBR",
           "avAdresseOWSTEXT",
           "avOppstartAarOWSNMBR"
       ],
        EnableInterleaving: true
    });
    
    console.log(results2.ElapsedTime);
    console.log(results2.RowCount);
    console.log(results2.PrimarySearchResults);

   

    if (results2) {
        context.res = {
            // status: 200, /* Defaults to 200 */ 
            body: results2.PrimarySearchResults
        };
    }
    else {
        context.res = {
            status: 400,
            body: "No contacts found"
        };
    }
};

export default getOpenListings;
