const express = require('express');
const axios = require('axios');
const qs = require('querystring');
const cors = require('cors');
require('dotenv').config();

const axiosInstance = axios.create({
    baseURL: 'https://login.microsoftonline.com/'
});

async function getToken(code, redirect_uri, tenant_id ){
    try{
        const body = qs.stringify({
            client_id: process.env.MICROSOFT_CLIENT_ID,
            grant_type: 'authorization_code',
            code,
            redirect_uri: redirect_uri,
            client_secret: process.env.MICROSOFT_CLIENT_SECRET,
            scope: 'offline_access user.read mail.read',
        });

        const options = {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        };

        const {data} = await axiosInstance.post(tenant_id+'/oauth2/v2.0/token', body, options);
        console.log(data);
        return data;

    }catch(err){
        return err;
    }
};

const server = express();

server.use(express.json());
server.use(cors());

server.post('/token', async (request,response)=>{
    const { code, redirect_uri, tenant_id  } = request.body; // suggestion-> scope can be passed here
    
    if(!code || !redirect_uri || !tenant_id)
    return response.status(400).json({error: 'the body must contain code, redirect_uri and tenant_id'});

    const token = await getToken(code, redirect_uri, tenant_id );
    
    if(token instanceof Error)
    return response.status(401).json({error:token.message});
    

    return response.json({token});
});



server.listen(process.env.PORT || 3003, ()=>{
    console.log(`server is on, port: ${process.env.PORT || 3003}`);
})