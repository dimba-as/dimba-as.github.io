

module.exports = async function getToken(context) {
    const adal = require('adal-node');
    return new Promise((resolve, reject) => {
  
      const authContext = new adal.AuthenticationContext("https://login.microsoftonline.com/b0c70ac8-59b9-4930-a342-3c6821f298c8");
  
      let secret = "BzLS6cexStyz57TBaL9--K.c.03WpTa8U_";
      let id = "8220e003-d7c3-48ad-aded-3ecddef6d5ed";
      
      authContext.acquireTokenWithClientCredentials("https://graph.microsoft.com/", id, secret, (err, tokenRes) => {
        if (err) {
          reject(err);
        } else {
          resolve(tokenRes.accessToken);
        }
      });
    });
  }