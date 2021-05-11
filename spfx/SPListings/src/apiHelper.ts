export const getSPListings = async()=>{
    const response = await fetch("https://dimbaproxy.azurewebsites.net/api/getSPListings?code=00Tmi1gESUAZL/nAfq74vEeF81xCJOoNaJko1MX1/dYyQbwHeaIVHg==");
    return response.json();
};
export const getOpenListings = async ()=>{
    const response = await fetch("https://dimbaproxy.azurewebsites.net/api/getOpenListings?code=4OqBaSOgsrDB9mPSS10FtgckI8GLbPMyqjaslWblYXIs4Jdx392nog==");
    return response.json();
};
