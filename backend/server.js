const express = require('express');
const cors = require('cors');
const app = express();
/**
 * APP ID === Client Id
 */
const APP_ID = "232181dd-d281-40f0-8c19-37234553bd57";
/**
 * Use one tannent we can invoke as many application
 */
const TENANT_ID="98d53658-2931-4186-ac56-42a6b796e78b";
/**
 * 
 */
const SECRET="TM1NA9Oy7FmvLje~.Wl0FavyRO1_-m0..l";
/**
 * 
 */
const SCOPE = "https://graph.microsoft.com/.default";

/**
 * 
 */

 const GRANT_TYPE="client_credentials";


/**
 * OBJECT_ID can be any user who is in application active direcory
 * 
 */
const OBJECT_ID="890d2bbc-5097-4688-97c4-742b5ebb5dcc"
let token="";

function getToken(){

}



app.get('/getMeetingLink',(req,res)=>{

    }
)

app.listen(8000,()=>{
    console.log("Server started on port 8000")
})