require('dotenv').config({
  path: require('path').join(__dirname, '.env')
})

const restify = require('restify')

const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  TeamsActivityHandler,
  MessageFactory
} = require('botbuilder')

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MICROSOFT_APP_ID,
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: process.env.MICROSOFT_APP_TYPE,
  MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID
})
const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory)
const adapter = new CloudAdapter(botFrameworkAuthentication)
adapter.onTurnError = (_context, error) => console.dir(error)

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
          await context.sendActivity(MessageFactory.text('login'))
          break;
        default:
          await context.sendActivity(MessageFactory.text('what?'))
          break;
      }
    })
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