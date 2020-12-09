import * as Express from "express";
import * as http from "http";
import * as path from "path";
import * as morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import * as debug from "debug";
import * as compression from "compression";



// Initialize debug logging module
const log = debug("msteams");

log(`Initializing Microsoft Teams Express hosted App...`);

// Initialize dotenv, to use .env file settings if existing
// tslint:disable-next-line:no-var-requires
require("dotenv").config();



// The import of components has to be done AFTER the dotenv config
import * as allComponents from "./TeamsAppsComponents";

// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;

// Inject the raw request body onto the request object
express.use(Express.json({
    verify: (req, res, buf: Buffer, encoding: string): void => {
        (req as any).rawBody = buf.toString();
    }
}));
express.use(Express.urlencoded({ extended: true }));

// Express configuration
express.set("views", path.join(__dirname, "/"));

// Add simple logging
express.use(morgan("tiny"));

// Add compression - uncomment to remove compression
express.use(compression());

// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsApiRouter(allComponents));

// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents
}));

// Set default web page
express.use("/", Express.static(path.join(__dirname, "web/"), {
    index: "index.html"
}));

// Set the port
express.set("port", port);


// Send get user role
express.get("/api/role", async (req, res, next) => {

    // console.log(req.query);

    const meetingId = req.query.meetingId;
    const userId = req.query.userId;

    const accessToken = await getAuthTokenFromMicrosoft(process.env.MICROSOFT_APP_ID, process.env.MICROSOFT_APP_PASSWORD);
    const memberRoleResponse = await getMeetingParticipant(accessToken, meetingId, userId, process.env.TENANT_ID as string);
    const userRole = memberRoleResponse.meetingRole;

    // console.log(meetingId);
    // console.log(userId);

    await res.json({ role: userRole });
    next();
});

// Sends the bubble notification
express.get("/api/bubble", async (req, res, next) => {

    // console.log(req.query);

    const chatId = req.query.chatId;
    const meetingBubbleTitle = "Contoso";

    const accessToken = await getAuthTokenFromMicrosoft(process.env.MICROSOFT_APP_ID, process.env.MICROSOFT_APP_PASSWORD);
    await sendBubbleMessage(accessToken, chatId, meetingBubbleTitle);

    res.sendStatus(200);
    next();
});


// Start the webserver
http.createServer(express).listen(port, () => {
    log(`Server running on ${port}`);
});

const getAuthTokenFromMicrosoft = async (appId, appSecret) => {

    const details = {
        scope: "https://api.botframework.com/.default",
        grant_type: "client_credentials",
        client_id: appId,
        client_secret: appSecret
    };

    const formBody: string[] = [];

    for (const property in details) {
      if (details.hasOwnProperty(property)) {
        const encodedKey = encodeURIComponent(property);
        const encodedValue = encodeURIComponent(details[property]);
        formBody.push(encodedKey + "=" + encodedValue);
      }
    }

    const postBody = formBody.join("&");

    const res = await fetch("https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: postBody,
    });
    const json = await res.json();
    if (json.error) { throw new Error(`${json.error}: ${json.error_description}`); }
    return json.access_token;
};

const getMeetingParticipant = async (token, meetingId, participantId, tenandId) => {

    // /v1/meetings/{meetingId}/participants/{participantId}?tenantId={tenantId}
    const res = await fetch(`https://smba.trafficmanager.net/amer/v1/meetings/${meetingId}/participants/${participantId}?tenantId=${tenandId}`, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      }
    });
    const json = await res.json();
    if (json.error) { throw new Error(`${json.error}: ${json.error_description}`); }
    return json;

};

const sendBubbleMessage = async (token, chatid, meetingBubbleTitle) => {

    // /v1/meetings/{meetingId}/participants/{participantId}?tenantId={tenantId}
    const res = await fetch(`https://smba.trafficmanager.net/amer/v3/conversations/${chatid}/activities`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({
        type: "message",
        text: "John Phillips assigned you a weekly todo",
        summary: "Don't forget to meet with Marketing next week",
        channelData: {
            notification: {
                alertInMeeting: true,
                externalResourceUrl: `https://teams.microsoft.com/l/bubble/${process.env.APPLICATION_ID}?url=${process.env.MEETING_TAB_URL}/&height=700&width=700&title=${meetingBubbleTitle}&completionBotId=${process.env.MICROSOFT_APP_ID}`
            }
        }
    })
    });
    const json = await res.json();
    if (json.error) { throw new Error(`${json.error}: ${json.error_description}`); }
    return json;

};
