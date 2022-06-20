import axios from 'axios'
const qs = require("qs")
require("dotenv").config()

// se importan las interfaces
import { GraphBody, item } from './interfaces/GraphBody'

const APP_ID = process.env.APP_ID+"";
const APP_SECRET = process.env.APP_SECRET+"";
const TOKEN_ENDPOINT = process.env.TOKEN_ENDPOINT+"";
const MS_GRAPH_SCOPE = process.env.MS_GRAPH_SCOPE+"";

const IdSitio = process.env.SITE_SHAREPOINT_ID+"";
const IdList = process.env.LIST_SHAREPOINT_ID+"";

// se obtiene el token para poder realizar las peticiones al graph
export const getAccessToken = async(graphBody: GraphBody, tokenEndPoint: string) => {
    const url = tokenEndPoint;
    try {      
        const response = await axios.post(url, qs.stringify(graphBody));
        if (response.status == 200) {
            return response.data.access_token;
        } else {
          throw new Error("Non 200OK response on obtaining token...");
        }
    } catch (error) {
        throw new Error("Error on obtaining token... "+error);
    }
}


// se crean registros en la lista de sharepoint
export const createItem = async (tokenGraph: string, body2: item) => {
  
    
    const url2 = `https://graph.microsoft.com/v1.0/sites/${IdSitio}/lists/${IdList}/items`            

    
    try {        
        const response = await axios({
            method: 'post',
            url: url2,           
            headers: {               
                "Authorization": `Bearer ${tokenGraph}`,
                "Content-Type": "application/json"
            },
            data: body2         
        });;                
               
        
    } catch (error) {
        throw new Error("Error on obtaining token... "+error);
    }
}

//obtener todos los items
export const getItems =async (tokenGraph: string) => {
    const url = `https://graph.microsoft.com/v1.0/sites/${IdSitio}/lists/${IdList}/items?expand=field`            

    try {        
        const response = await axios.get(url, { headers: { Authorization: tokenGraph }});;                
                
        const items = response.data.value
        const items2 = items.map(el => ( 
            { 
                ID: el.id,
                Title: el.fields.Title
            } 
        ));;
        console.log("--------Todos los items---------");
        
        console.log(items2);              

    } catch (error) {
        throw new Error("Error on obtaining token... " + error);
    }
}

//obtener item individual
export const getItem =async (idItem: string, tokenGraph: string) => {    

    const url = `https://graph.microsoft.com/v1.0/sites/${IdSitio}/lists/${IdList}/items/${idItem}?expand=fields`            

    try {        
        const response = await axios.get(url, { headers: { Authorization: tokenGraph }});;                

        const items = response.data

        const item = { 
            ID: response.data.id,
            Title: response.data.fields.Title
        } 
        console.log("--------Item individual---------");
        console.log(item);              

    } catch (error) {
        throw new Error("Error on obtaining token... " + error);
    }
}