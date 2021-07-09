import axios from 'axios';
export const getSPListings = async () => {
    const response = await fetch("https://dimbaproxy.azurewebsites.net/api/getSPListings?code=00Tmi1gESUAZL/nAfq74vEeF81xCJOoNaJko1MX1/dYyQbwHeaIVHg==");
    return response.json();
};
export const getOpenListings = async () => {
    const response = await fetch("https://dimbaproxy.azurewebsites.net/api/getOpenListings?code=4OqBaSOgsrDB9mPSS10FtgckI8GLbPMyqjaslWblYXIs4Jdx392nog==");
    return response.json();
};
export const flagInterest = async (props) => {
    const url = "https://dimbaproxy.azurewebsites.net/api/flagInterest?code=nM9MS3VOxWEcJTFD3RPFovDBDHLswBdwX/06h2Pb1uTgiiIYqX1gVQ==";
    const response = await axios.post(url,props);
    // const response = await fetch(url,
    //     {
    //         method: "POST",
    //         body: props,
    //         headers: {
    //             "Content-type": "application/json; charset=UTF-8"
    //         }
    //     });
    const result = response.status;
    return result;
};