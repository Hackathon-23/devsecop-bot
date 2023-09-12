import { BotDeclaration } from "express-msteams-host";
import * as debug from "debug";
import {
    CardFactory,
    ConversationState,
    MemoryStorage,
    UserState,
    TurnContext,
    teamsGetChannelId,
    BotFrameworkAdapter,
    Activity,
    ConversationParameters,
    ActionTypes,
    ChannelInfo,
    ConversationReference,
    ConversationResourceResponse,
    TeamsChannelData,
    MessageFactory
} from "botbuilder";
import { DialogBot } from "./dialogBot";
import { MainDialog } from "./dialogs/mainDialog";
import WelcomeCard from "./cards/welcomeCard";

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for DevSecOps Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_ID,
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_PASSWORD)

export class DevSecOpsBot extends DialogBot {
    constructor(conversationState: ConversationState, userState: UserState) {
        super(conversationState, userState, new MainDialog());

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            if (membersAdded && membersAdded.length > 0) {
                for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                    if (membersAdded[cnt].id !== context.activity.recipient.id) {
                        const customerId = await context.activity.recipient.id;
                        console.log("This is the customer id: ", customerId);
                        await this.sendWelcomeCard(context);
                    }
                }
            }
            await next();
        });

        this.onMessage(async (context: TurnContext, next: () => Promise<void>) => {
            const botMessageText: string = context.activity.text.trim().toLowerCase();
            if (botMessageText.startsWith("code")) {
                const channelId = await context.activity.conversation.id;
                console.log("This is the channel id: " + channelId);
                const message = MessageFactory.text("This will be the first message in a new thread");
                let text = TurnContext.removeRecipientMention(context.activity);
                text = text.toLowerCase();
                const newConversation = await this.createConversationInChannel(context, channelId, message);
                // const newConversation = await this.createConversationInChannel(context, channelId, message);
            }
            await next();
        });
    }

    public async sendWelcomeCard(context: TurnContext): Promise<void> {
        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
        await context.sendActivity({ attachments: [welcomeCard] });
    }

    private async teamsCreateConversation(context: TurnContext, message: Partial<Activity>): Promise<void> {
        // get a reference to the bot adapter and create a connection to the Teams API
        const adapter = <BotFrameworkAdapter>context.adapter;
        const connectorClient = adapter.createConnectorClient(context.activity.serviceUrl);
        // set current teams channel in new conversation parameters
        const teamsChannelId = teamsGetChannelId(context.activity);
        const conversationParameters: ConversationParameters = {
            isGroup: true,
            channelData: {
                channel: {
                    id: teamsChannelId
                }
            },
            activity: message as Activity,
            bot: context.activity.recipient
        };

        // create conversation and send message
        await connectorClient.conversations.createConversation(conversationParameters);
    }

    private async createConversationInChannel(context: TurnContext, teamsChannelId: string, message: Partial<Activity>): Promise<[ConversationReference, string]> {
        // create parameters for the new conversation
        const conversationParameters = <ConversationParameters>{
            isGroup: true,
            channelData: <TeamsChannelData>{
                channel: <ChannelInfo>{
                    id: teamsChannelId
                }
            },
            activity: message
        };

        // get a reference to the bot adapter and create a connection to the Teams API
        const adapter = <BotFrameworkAdapter>context.adapter;
        const connectorClient = adapter.createConnectorClient(context.activity.serviceUrl);

        // create a new conversation and get a reference to it
        const conversationResourceResponse: ConversationResourceResponse = await connectorClient.conversations.createConversation(conversationParameters);
        const conversationReference = <ConversationReference>TurnContext.getConversationReference(context.activity);
        conversationReference.conversation.id = conversationResourceResponse.id;

        return [conversationReference, conversationResourceResponse.activityId];
    }
}
