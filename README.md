# sso demo
- https://github.com/OfficeDev/TrainingContent/tree/master/Teams/80%20Using%20Single%20Sign-On%20with%20Microsoft%20Teams/Demos/02-learn-msteams-sso-bot
- https://youtu.be/cmI06T2JLEg
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/24.bot-authentication-msgraph

```
 await (turnContext.adapter as BotFrameworkAdapter).exchangeToken(
        turnContext,
        tokenExchangeRequest.connectionName,
        turnContext.activity.from.id,
        tokenExchangeRequest)
```