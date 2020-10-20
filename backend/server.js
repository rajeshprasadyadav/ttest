const express = require('express');
const cors = require('cors');
const app = express();
const qs = require('qs');
const axios = require('axios');
const bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: true }));
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
 * OBJECT_ID can be any user who is in application active direcory
 * It is basically an User wpplication will create an online event on behalf of that.
 * 
 */
const OBJECT_ID="890d2bbc-5097-4688-97c4-742b5ebb5dcc";

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
   * JOIN Link URL
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