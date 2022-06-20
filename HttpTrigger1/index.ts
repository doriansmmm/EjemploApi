import { AzureFunction, Context, HttpRequest } from "@azure/functions"
require("dotenv").config()

// se importan las funciones
import { 
    getAccessToken,
    createItem,
    getItems,
    getItem
} from './graph'

// se importan las interfaces
import { item } from './interfaces/GraphBody'

const APP_ID = process.env.APP_ID+"";
const APP_SECRET = process.env.APP_SECRET+"";
const TOKEN_ENDPOINT = process.env.TOKEN_ENDPOINT+"";
const MS_GRAPH_SCOPE = process.env.MS_GRAPH_SCOPE+"";

const body = {
    client_id: APP_ID,
    scope: MS_GRAPH_SCOPE,
    client_secret: APP_SECRET,
    grant_type: "client_credentials",
};


const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const name = (req.query.name || (req.body && req.body.name));
    let accesToken: any
    let ctItem: any
    let gItems: any
    let gItem: any
    try{       
        
        accesToken = await getAccessToken(body, TOKEN_ENDPOINT)
        const body2: item = {
            "fields": {
                "Title": "registro2"
            }  
        }

        const idItem = "40"
        ctItem = await createItem(accesToken, body2)
        gItems = await getItems(accesToken)
        gItem = await getItem(idItem, accesToken)
    }catch{}




    context.res = {
        // status: 200, /* Defaults to 200 */
        body: accesToken
    };

};

export default httpTrigger;