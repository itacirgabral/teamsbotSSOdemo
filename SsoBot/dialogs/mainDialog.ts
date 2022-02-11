import {
  StatePropertyAccessor,
  TurnContext
} from "botbuilder-core";
import {
  DialogSet,
  DialogState,
  DialogTurnResult,
  DialogTurnStatus,
  WaterfallDialog,
  WaterfallStepContext
} from "botbuilder-dialogs";
import { LogoutDialog } from "./logoutDialog";
import { SsoOauthPrompt } from "./ssoOauthPrompt";
import "isomorphic-fetch";
import { MsGraphHelper } from "../helpers/MsGraphHelper";

const MAIN_DIALOG_ID = "MainDialog";
const MAIN_WATERFALL_DIALOG_ID = "MainWaterfallDialog";
const OAUTH_PROMPT_ID = "OAuthPrompt";

export class MainDialog extends LogoutDialog {
  constructor() {
    super(MAIN_DIALOG_ID, process.env.SSO_CONNECTION_NAME as string);

    console.log("MainDialog constructor");

    const ssoOauthPrompt = new SsoOauthPrompt(OAUTH_PROMPT_ID, {
      connectionName: process.env.SSO_CONNECTION_NAME as string,
      text: "Please sign in",
      title: "Sign In",
      timeout: 300000
    })

    console.log("###########################################################################################")
    console.dir({ ssoOauthPrompt })
    console.log(JSON.stringify(ssoOauthPrompt, null, 2))

    this.addDialog(ssoOauthPrompt);

    const waterfallDialog = new WaterfallDialog(MAIN_WATERFALL_DIALOG_ID, [
      this.promptStep.bind(this),
      this.displayMicrosoftGraphDataStep.bind(this)
    ])
    // console.log(JSON.stringify(waterfallDialog, null, 2))
    this.addDialog(waterfallDialog);

    // set the initial dialog to the waterfall
    this.initialDialogId = MAIN_WATERFALL_DIALOG_ID;
  }

  public async run(turnContext: TurnContext, accessor: StatePropertyAccessor<DialogState>): Promise<void> {
    const dialogSet = new DialogSet(accessor);
    dialogSet.add(this);
    const dialogContext = await dialogSet.createContext(turnContext);
    const results = await dialogContext.continueDialog();
    if (results.status === DialogTurnStatus.empty) {
      console.log('run 0')
      await dialogContext.beginDialog(this.id);
      console.log('run f')
    }
  }

  public async promptStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    try {
      console.log('promptStep 0')
      const dialog = await stepContext.beginDialog(OAUTH_PROMPT_ID);
      console.log('promptStep f')
      return dialog
    } catch (err) {
      console.error(err);
    }
    return await stepContext.endDialog();
  }

  public async displayMicrosoftGraphDataStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    console.log("displayMicrosoftGraphDataStep");
    // get token from prev step (or directly from the prompt itself)
    const tokenResponse = stepContext.result;
    if (!tokenResponse?.token) {
      await stepContext.context.sendActivity("Login not successful, please try again.");
    } else {

      console.dir(stepContext)
      const msGraphClient = new MsGraphHelper(tokenResponse?.token);

      const user = await msGraphClient.getCurrentUser();
      await stepContext.context.sendActivity(`Thank you for signing in ${user.displayName as string} (${user.userPrincipalName as string})!`);
      await stepContext.context.sendActivity("I can retrieve your details from Microsoft Graph using my support for SSO!");

      const email = await msGraphClient.getMostRecentEmail();
      await stepContext.context.sendActivity(`Your most recent email about "${email.subject as string}" was received at ${new Date(email.receivedDateTime as string).toLocaleString()}.`);
    }

    return await stepContext.endDialog();
  }
}
