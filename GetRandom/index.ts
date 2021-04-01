import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');

const APP_ID = process.env["APP_ID"];
const APP_SECRET = process.env["APP_SECRET"];
const TENANT_ID = process.env["TENANT_ID"];
const SITE_ID = process.env["SITE_ID"];
const LIST_ID = process.env["LIST_ID"];


const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/';


const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    
    // Set Default Header for Axios Requests
    axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

    // Get Token for MS Graph
    let token = await getToken();

    let listitems = await getSPListItems(token);

    let returnValue = "";

    for(let item of listitems) {
        returnValue += item.fields.Title + " - " + item.fields.DemoText + "\n"
    }

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: returnValue
    };

}
export default httpTrigger;

/**
 * Get Token for MS Graph
 */
async function getToken(): Promise<string> {
    const postData = {
        client_id: APP_ID,
        scope: MS_GRAPH_SCOPE,
        client_secret: APP_SECRET,
        grant_type: 'client_credentials'
    };

    return await axios
        .post(TOKEN_ENDPOINT, qs.stringify(postData))
        .then(response => {
            // console.log(response.data);
            return response.data.access_token;
        })
        .catch(error => {
            console.log(error);
        });
}

/**
 * Get SP Sites
 * @param token Token to authenticate through MS Graph
 */
async function getSPListItems(token:string): Promise<ListItem[]> {
    let config: AxiosRequestConfig = {
        method: 'get',
        url: MS_GRAPH_ENDPOINT + 'sites/' + SITE_ID + '/lists/' + LIST_ID + '/items?expand=fields',
        headers: {
          'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        }
    }

    console.log(MS_GRAPH_ENDPOINT + 'sites/' + SITE_ID + '/lists/' + LIST_ID + '/items?expand=fields');
    
    return await axios(config)
        .then(response => {
            console.log(response.data);
            return response.data.value;
        })
        .catch(error => {
            console.log(error);
        });
}

class ListItem {
    createdDateTime:        Date;
    eTag:                   string;
    id:                     string;
    lastModifiedDateTime:   Date;
    webUrl:                 string;
    createdBy:              UserObject;
    lastModifiedBy:         UserObject;
    parentReference:        ParentReference;
    contentType:            ContentType;
    "fields@odata.context": string;
    fields:                 Fields;
}

class ContentType {
    id:   string;
    name: string;
}

class UserObject {
    user: User;
}

class User {
    email:       string;
    id:          string;
    displayName: string;
}

class Fields {
    "@odata.etag":             string;
    id:                        string;
    ContentType:               string;
    Title:                     string;
    Modified:                  Date;
    Created:                   Date;
    AuthorLookupId:            string;
    EditorLookupId:            string;
    DemoText:                  string;
}

class ParentReference {
    siteId: string;
}