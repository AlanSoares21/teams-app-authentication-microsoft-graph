import axios from 'axios';

const request = axios.create({
    baseURL: process.env.TOKEN_URL || 'http://localhost:3003'
});

export async function getToken(code: string, redirect_uri: string, tenant_id: string){
    try{
        const body = {code, redirect_uri, tenant_id};
        const {data} = await request.post('token', body);
        return data;
    }catch(error){
        console.error(error);
        return error;
    }
    
}