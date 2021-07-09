const SPFetchClient = require("@pnp/nodejs-commonjs").SPFetchClient;
const sp = require("@pnp/sp-commonjs").sp;

module.exports = async function connectSPO(url:string){
    if(url.indexOf("/sites")>0){
        url = "https://dimbaas.sharepoint.com/sites/aviadoas";
    }
   
   return sp.setup({
        sp: {
            fetchClientFactory: () => {
                //elevate
                //return new SPFetchClient(url,"8c92bf25-bf7c-41b9-a7bb-adbd7c6cd93e", "3PVAa~6ZU-qz3tqwML4x4-I6bp1zU.XofJ");
                //aviado
                //return new SPFetchClient(url,"38f9310a-0029-4692-9ca5-13e661511665", "LI-8s8ymcAhBGHPfqT.Rr~yFl_7i-DV2CM");
                //dimba
                return new SPFetchClient(url,"8220e003-d7c3-48ad-aded-3ecddef6d5ed", "BzLS6cexStyz57TBaL9--K.c.03WpTa8U_");
            },
        },
    });
}