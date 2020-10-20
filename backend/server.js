const express = require('express');
const cors = require('cors');
const app = express();
const qs = require('qs');
const axios = require('axios');
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
 * constant throughout
 */
const SCOPE = "https://graph.microsoft.com/.default";

/**
 * 
 * Const througout
 */

 const GRANT_TYPE="client_credentials";

 /**
  *  Token URL 
  */

  const TOKEN_URL = "https://login.microsoftonline.com/"+TENANT_ID+"/oauth2/v2.0/token";

/**
 * OBJECT_ID can be any user who is in application active direcory
 * 
 */
const OBJECT_ID="890d2bbc-5097-4688-97c4-742b5ebb5dcc"
let token="";

function getToken(){
    
}


const tokenRequestBody=qs.stringify({
    client_id:APP_ID,
    scope:SCOPE,
    client_secret:SECRET,
    grant_type:GRANT_TYPE
})
const tokenRequestHeader = {
    'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
}

app.get('/getMeetingLink',(req,res)=>{
    axios.post(TOKEN_URL,
        { client_id:APP_ID,
            scope:SCOPE,
            client_secret:SECRET,
            grant_type:GRANT_TYPE
        },{
            headers:{
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        }
        ).then(result=>{
            console.log(result)
        });
}
)

app.get("/test",(req,res)=>{
    res.status(200).json({"message":"request successfull"});
})

app.listen(8000,()=>{
    console.log("Server started on port 8000")
})