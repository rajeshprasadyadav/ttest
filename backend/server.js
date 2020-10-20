const express = require('express');
const cors = require('cors');
const app = express();
const qs = require('qs');
const axios = require('axios');
const bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: true }));

/**
 * APP ID,Unique for each application, we create in (portal.azure.com);
 * We can co-realte it with each partner/cohort 
 * at run time these credentials will be pulled from database
 */
const APP_ID = "232181dd-d281-40f0-8c19-37234553bd57";
/**
 * Unique for each Microsoft account,basically unique for service provider
 * Tenant_id is required for getting token,irrespective of APP_ID.
 */
const TENANT_ID="98d53658-2931-4186-ac56-42a6b796e78b";
/**
 * Secret is required to get token,each application will have secret key as well
 */
const SECRET="TM1NA9Oy7FmvLje~.Wl0FavyRO1_-m0..l";
/**
 * this is constant and required for getting token
 */
const SCOPE = "https://graph.microsoft.com/.default";



/**
 * Object ID , represent any activie user under Azure Active Directory
 *
 */
const OBJECT_ID="890d2bbc-5097-4688-97c4-742b5ebb5dcc";

//another user id === ObjectID
//const OBJECT_ID="f89fdfa5-bec2-421b-a9ef-b0016f5bcb09";

/**
 * 
 * Constant required for get access_token
 */

const GRANT_TYPE="client_credentials";

/**
 * End point for getting access_token,
 * Once We will have access token we can hit the end for for Creating Event.
 * Required Field: app_id,secret,scope & grant_type , (last two fields is constant)
 */
const TOKEN_URL = "https://login.microsoftonline.com/"+TENANT_ID+"/oauth2/v2.0/token";

/**
 * 
 */
const JOIN_LINK_URL="https://graph.microsoft.com/v1.0/users/"+OBJECT_ID+"/calendar/events";

let token="";

async function getToken(){
    try{
    return await axios.post(TOKEN_URL,
        qs.stringify({
            'client_id':APP_ID,
            'scope':SCOPE,
            'client_secret':SECRET,
            'grant_type':GRANT_TYPE
        }),{
            headers:{
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        }
        )
    }catch(err){
        console.log(err);
    }
}
async function getJoinLink(token,params){
    try{
        return await axios.post(
            JOIN_LINK_URL,
            {
                "subject":params.subject,
                "body": {
                    "contentType": "HTML",
                    "content": "Does noon work for you?"
                },
                "start": {
                    "dateTime": params.startDateAndTime,
                    "timeZone": params.timeZone
                },
                "end": {
                    "dateTime": params.endDateAndTime,
                    "timeZone": params.timeZone
                },
                "location":{
                    "displayName":"Coaching Session"
                },
                "attendees": [
                    {
                    "emailAddress": {
                    "address" :params.attendeeEmail,
                    "name": params.attendeeName
                    },
                    "type": "required"
                }
                ],
                "allowNewTimeProposals": true,
                "isOnlineMeeting": true,
                "onlineMeetingProvider": "teamsForBusiness"
            },
                {
                    headers:{
                        'Authorization':"Bearer "+token,
                        'Content-Type':'application/json'
                }
            }
        )
    }catch(err){
        console.log("axios error ++++++++++");
        console.log(err);
    }

}


app.post('/getMeetingLink',(req,res)=>{
   getToken()
    .then(tokenResponse=>{
        if(tokenResponse.status===200){
            token = tokenResponse.data.access_token;
            let params = {
                subject:req.body.subject,
                startDateAndTime:req.body.startDateAndTime,
                endDateAndTime:req.body.endDateAndTime,
                timeZone:req.body.timeZone,
                attendeeEmail:req.body.attendeeEmail,
                attendeeName:req.body.attendeeName
            }
            getJoinLink(token,params).then(meetingResponse=>{
                if(meetingResponse.status===201){
                    res.status(201).json(meetingResponse.data);
                    console.log('Created Succesfully');
                    
                }else if(res.status===401){
                    res.status(401).json({error:"Unauthorized request"})
                }

               
            }).catch(err=>{
                console.log("Request error =============");
                console.log(err)
            })
        }
    })
  
   
}
)
app.listen(8000,()=>{
    console.log("Server started on port 8000")
})