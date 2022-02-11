import { BotDeclaration } from "express-msteams-host";
import * as debug from "debug";
import {
  ConversationState,
  UserState,
  SigninStateVerificationQuery,
  TurnContext,
  MemoryStorage
} from "botbuilder";
import { DialogBot } from "./DialogBot";
import { MainDialog } from "./dialogs/mainDialog";
import { SsoOAuthHelper } from "./helpers/SsoOauthHelper";

// Initialize debug logging module
const log = debug("msteams");

@BotDeclaration(
  "/api/messages",
  new MemoryStorage(),
  // eslint-disable-next-line no-undef
  process.env.MICROSOFT_APP_ID,
  // eslint-disable-next-line no-undef
  process.env.MICROSOFT_APP_PASSWORD)
export class SsoBot extends DialogBot {
  public _ssoOAuthHelper: SsoOAuthHelper;

  constructor(conversationState: ConversationState, userState: UserState) {
    super(conversationState, userState, new MainDialog());
    this._ssoOAuthHelper = new SsoOAuthHelper();
    // ssoOAuthHelper.shouldProcessTokenExchange
    // ssoOAuthHelper.exchangeToken

    console.log("SsoBot constructor");

    this.onMembersAdded(async (context, next) => {
      console.log("onMembersAdded");
      const membersAdded = context.activity.membersAdded;
      if (membersAdded && membersAdded.length > 0) {
        for (let cnt = 0; cnt < membersAdded.length; cnt++) {
          if (membersAdded[cnt].id !== context.activity.recipient.id) {
            await context.sendActivity("Welcome to TeamsBot. Type anything to get logged in. Type 'logout' to sign-out.");
          }
        }
      }
      await next();
    });

    this.onTokenResponseEvent(async (context, next) => {
      console.log("SsoBot onTokenResponseEvent");
      await this.dialog.run(context, this.dialogState);
      await next();
    });
  }

  public async handleTeamsSigninTokenExchange(context: TurnContext, query: SigninStateVerificationQuery): Promise<void> {
    console.log("SsoBot handleTeamsSigninTokenExchange");
    if (!await this._ssoOAuthHelper.shouldProcessTokenExchange(context)) {
      await this.dialog.run(context, this.dialogState);
    }
  }

  public async handleTeamsSigninVerifyState(context: TurnContext, query: SigninStateVerificationQuery): Promise<void> {
    console.log("SsoBot handleTeamsSigninVerifyState");
    await this.dialog.run(context, this.dialogState);
  }
}
