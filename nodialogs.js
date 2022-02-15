require('dotenv').config({
  path: require('path').join(__dirname, '.env')
})

const SsoConnectionName = process.env.SSO_CONNECTION_NAME
const MicrosoftAppId = process.env.MICROSOFT_APP_ID
const MicrosoftAppPassword = process.env.MICROSOFT_APP_PASSWORD
const MicrosoftAppType = process.env.MICROSOFT_APP_TYPE
const MicrosoftAppTenantId = process.env.MICROSOFT_APP_TENANT_ID

const restify = require('restify')

const {
  CardFactory,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  TeamsActivityHandler,
  tokenExchangeOperationName,
  MessageFactory
} = require('botbuilder')


const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId,
  MicrosoftAppPassword,
  MicrosoftAppType,
  MicrosoftAppTenantId
})
const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory)
const adapter = new CloudAdapter(botFrameworkAuthentication)
adapter.onTurnError = (_context, error) => console.dir(error)

let token
let msGraphClient
const microsoft = require('@microsoft/microsoft-graph-client')
require('isomorphic-fetch')

const TeamsConversationBot = class TeamsConversationBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onConversationUpdate(async (context, next) => {
      console.log('onConversationUpdate')
    })

    this.onInstallationUpdate(async (context, next) => {
      console.log('onInstallationUpdate')
    })

    this.onMessage(async (context, next) => {
      const text = context.activity?.text?.trim() ?? ''
      switch (text) {
        case 'login':
          if (token) {
            await context.sendActivity(MessageFactory.text(token))
          } else {
          const oauthCard = await CardFactory.oauthCard(SsoConnectionName, undefined, undefined, undefined, {
            id: 'random65jHf9276hDy47',
            uri: `api://botid-${MicrosoftAppId}`
          })
          await context.sendActivity(MessageFactory.attachment(oauthCard))
          await context.sendActivity(MessageFactory.text('login'))
          }
          break;
        case 'getme':
          if (msGraphClient) {
            const me = await msGraphClient.api("me").get()
            console.dir({ me })
            await context.sendActivity(MessageFactory.text('me'))
          } else {
            await context.sendActivity(MessageFactory.text('do login'))
          }
          break;
        default:
          await context.sendActivity(MessageFactory.text('what?'))
          break;
      }
    })
  }//

  async handleTeamsSigninTokenExchange(context, query) {
    console.log('handleTeamsSigninTokenExchange')
    if (context?.activity?.name === tokenExchangeOperationName) {
      token = context?.activity?.value?.token

      msGraphClient = microsoft.Client.init({
        debugLogging: true,
        authProvider: done => {
          done(null, token)
        }
      })
    }
  }
}
const bot = new TeamsConversationBot()

const server = restify.createServer()
server.use(restify.plugins.bodyParser())
server.post('/api/messages', async (req, res) => {
  console.log('POST /api/messages')
  console.dir(req.body)
  await adapter.process(req, res, context => bot.run(context))
})

server.listen(process.env.port || process.env.PORT || 3978, function() {
  console.log(`\n${ server.name } listening to ${ server.url }`);
})