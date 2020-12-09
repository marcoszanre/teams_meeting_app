import { BotDeclaration, PreventIframe, BotCallingWebhook, MessageExtensionDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, TeamsInfo, BotFrameworkAdapter, MessageFactory } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import MeetingAppMessageExtension from "../meetingAppMessageExtension/MeetingAppMessageExtension";
import WelcomeCard from "./dialogs/WelcomeDialog";
import express = require("express");

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for Meeting App Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)
@PreventIframe("/meetingAppBot/aboutMe.html")
export class MeetingAppBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    /** Local property for MeetingAppMessageExtension */
    @MessageExtensionDeclaration("meetingAppMessageExtension")
    private _meetingAppMessageExtension: MeetingAppMessageExtension;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();
        // Message extension MeetingAppMessageExtension
        this._meetingAppMessageExtension = new MeetingAppMessageExtension();


        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        // Set up the Activity processing

        this.onMessage(async (context: TurnContext): Promise<void> => {
            // TODO: add your own bot logic in here

            // const senderRole = await context.adapter as BotFrameworkAdapter;
            // const meuConnector = await senderRole.createConnectorClient("https://smba.trafficmanager.net/amer");
            
            // Convertsation Member
            console.log(context.activity);
            // this.returnConversationMember(context);

            const meetingTabURL = "https://5930ada58187.ngrok.io/meetingAppTab/index.html";
            const meetingBubbleTitle = "Contoso";

            const replyActivity = MessageFactory.text('Hi');
            replyActivity.channelData = {
                notification: {
                    alertInMeeting: true,
                    externalResourceUrl: `https://teams.microsoft.com/l/bubble/${process.env.APPLICATION_ID}?url=${meetingTabURL}/&height=300&width=400&title=${meetingBubbleTitle}&completionBotId=${process.env.MICROSOFT_APP_ID}`
                }
            };
            await context.sendActivity(replyActivity);

            switch (context.activity.type) {
                case ActivityTypes.Message:
                    let text = TurnContext.removeRecipientMention(context.activity);
                    text = text.toLowerCase();
                    if (text.startsWith("hello")) {
                        await context.sendActivity("Oh, hello to you as well!");
                        return;
                    } else if (text.startsWith("help")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog("help");
                    } else {
                        await context.sendActivity(`I\'m terribly sorry, but my developer hasn\'t trained me to do anything yet...`);
                    }
                    break;
                default:
                    break;
            }
            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });

    }

    public returnConversationMember = async (context: TurnContext) => {

        const accessToken = await this.getAuthTokenFromMicrosoft(process.env.MICROSOFT_APP_ID, process.env.MICROSOFT_APP_PASSWORD);
        console.log(accessToken);
        const memberRoleResponse = await this.getMeetingParticipant(accessToken, process.env.MEETING_ID as string, context.activity.from.aadObjectId, process.env.TENANT_ID as string);
        console.log(memberRoleResponse);
        return;

    }

    public getAuthTokenFromMicrosoft = async (appId, appSecret) => {

        const details = {
            scope: "https://api.botframework.com/.default",
            grant_type: "client_credentials",
            client_id: appId,
            client_secret: appSecret
        };
        
        const formBody: string[] = [];
        
        for (const property in details) {
          const encodedKey = encodeURIComponent(property);
          const encodedValue = encodeURIComponent(details[property]);
          formBody.push(encodedKey + "=" + encodedValue);
        }
        
        const postBody = formBody.join("&");

        const res = await fetch('https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
          body: postBody,
        });
        const json = await res.json();
        if (json.error) throw new Error(`${json.error}: ${json.error_description}`);
        return json.access_token;
    };

    public getMeetingParticipant = async (token, meetingId, participantId, tenandId) => {

        // /v1/meetings/{meetingId}/participants/{participantId}?tenantId={tenantId}
        const res = await fetch(`https://smba.trafficmanager.net/amer/v1/meetings/${meetingId}/participants/${participantId}?tenantId=${tenandId}`, {
          method: "GET",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${token}`
          }
        });
        const json = await res.json();
        if (json.error) throw new Error(`${json.error}: ${json.error_description}`);
        return json;

    }


    /**
     * Webhook for incoming calls
     */
    @BotCallingWebhook("/api/calling")
    public async onIncomingCall(req: express.Request, res: express.Response) {
        log("Incoming call");
        // TODO: Implement authorization header validation

        // TODO: Add your management of calls (answer, reject etc.)

        // default, send an access denied
        res.sendStatus(401);
    }
}
