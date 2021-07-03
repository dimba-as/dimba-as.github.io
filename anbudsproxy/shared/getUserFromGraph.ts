import * as request from 'request';
module.exports = async function getUserFromGraph(context, token, user) {
    return new Promise((resolve, reject) => {

        const url = `https://graph.microsoft.com/v1.0/users/?$filter=mail eq '${user}'&$select=CompanyName,id`;
        try {

            request.get(url, {
                "auth": {
                    "bearer": token
                },
                headers: {
                    "Content-Type": "application/json; charset=utf-8"
                }
                
            }, (error, response, body) => {

                if (!error) {
                    resolve(JSON.parse(body));
                    return;
                } else { 
                    reject(`Error in getChannelId: ${error}`);
                }

            });

        }
        catch (ex) {
            context.log(`Error in getChannelId: ${ex}`);
        }

    });
}