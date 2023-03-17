
const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const rawImageCard = require("./adaptiveCards/image.json");
const rawChatCard = require("./adaptiveCards/chat.json");
const rawChatResponseCard = require("./adaptiveCards/chatResponse.json")
const rawWeatherCard = require("./adaptiveCards/weather.json");
const rawWeatherResponseCard=require("./adaptiveCards/weatherResponse.json")
const cardTools = require("@microsoft/adaptivecards-tools");
const Axios = require("axios")
class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();



    

    // record the likeCount
    this.likeCountObj = { likeCount: 0 };
    this.weatherData={"city":"",temperature:23,humidity:34}
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        
        case "weather": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawWeatherCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }

        case "chat": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawChatCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }

        case "image": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawImageCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "hello": {
          
          // const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity("Hello world");
          break;
        }
    
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

  
  }


  async onAdaptiveCardInvoke(context, invokeValue) {

    if (invokeValue.action.verb === "weatherQuery") {
      console.log(invokeValue.action.data.weatherQuery)
      this.weatherData.city=invokeValue.action.data.weatherQuery
      let result = await Axios.get(`https://api.openweathermap.org/data/2.5/weather?q=${this.weatherData.city}&appid=ba5e4daae505c48512f7b1b07df8781f`);
      this.weatherData.temperature=result.data.main.temp 
      this.weatherData.humidity=result.data.main.humidity
      this.weatherData.pressure=result.data.main.pressure
      
      const weatherCard = cardTools.AdaptiveCards.declare(rawWeatherResponseCard).render(this.weatherData);
      await context.sendActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(weatherCard)],
      });
      return { statusCode: 200 };
    }
    else if(invokeValue.action.verb === "chatQuery"){
      const chatText={text:invokeValue.action.data.chatQuery}
      const chatCard = cardTools.AdaptiveCards.declare(rawChatResponseCard).render(chatText);
      await context.sendActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(chatCard)],
      });
      return { statusCode: 200 };
    }
  }

 
 
}



module.exports.TeamsBot = TeamsBot;
