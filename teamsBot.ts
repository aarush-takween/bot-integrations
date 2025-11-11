import {
  TeamsActivityHandler,
  TurnContext,
} from "botbuilder";
import * as https from "https";
import axios, { AxiosResponse } from 'axios';
import settings from './appSettings';
import * as graphHelper from './graphHelper';

export class TeamsBot extends TeamsActivityHandler {
  storage: any;
  constructor(dbStorage) {
    super();
    this.storage = dbStorage;
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();

      const userAdId = context.activity.from.aadObjectId;
      graphHelper.initializeGraphForAppOnlyAuth(settings);
      const userDetails = await graphHelper.getUserAsync(userAdId);
      await context.sendActivity("Your answer is coming right up!");
      let history = await callFromDb(this.storage, userAdId);
      let answer = await getGenAItext(txt, history, userDetails["mail"]);
      putToDb(this.storage, txt, answer, userAdId);
      context.activity.textFormat = "markdown";
      await context.sendActivity(answer);
      await next();
    });
  }
}

async function getGenAItext(query: string, history: Array<string>, userMail: string): Promise<string> {
  let url = process.env.URL;
  let token = process.env.TOKEN;
  let body = {
    "data": query,
    "history": history,
    "user": userMail
  };
  let response: AxiosResponse = await axios.post(url, body, {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': token
    },
    httpsAgent: new https.Agent({
      rejectUnauthorized: false
    })
  });
  return response.data;
}

async function putToDb(storage, question: string, answer: string, userAdId: string) {
  try {
    console.log("Inserting to DB!");
    let storedItems = await storage.read([userAdId]);
    var UserMessageData = storedItems[userAdId];
    if (typeof (UserMessageData) != 'undefined') {
      let messageLength = storedItems[userAdId].numMessage;
      if (messageLength < 5) {
        storedItems[userAdId].questionList.push(question);
        storedItems[userAdId].answerList.push(answer);
        storedItems[userAdId].numMessage++;
      }
      else {
        storedItems[userAdId].questionList.splice(0, 1);
        storedItems[userAdId].answerList.splice(0, 1);
        storedItems[userAdId].questionList.push(question);
        storedItems[userAdId].answerList.push(answer);
      }
      try {
        await storage.write(storedItems);
      } catch (err) {
          console.log(`Write failed: ${err}`);
      }
    }
    else {
      const numMessage = 1;
      storedItems[userAdId] = { questionList: [`${question}`], answerList: [`${answer}`], "eTag": "*", numMessage };
      try {
        await storage.write(storedItems);
      } catch (err) {
          console.log(`Write failed: ${err}`);
      }
    }
  }
  catch (err){
    console.log(`Read rejected. ${err}`);
  }
}

async function callFromDb(storage, userAdId: string) {
  try {
    console.log("Calling From DB!");
    let storedItems = await storage.read([userAdId]);
    var UserMessageData = storedItems[userAdId];
    if (typeof (UserMessageData) != 'undefined') {
      let messageLength = storedItems[userAdId].numMessage;
      let resultArray = [];
      for ( let i =0; i < messageLength; i++) {
        let question = storedItems[userAdId].questionList[i];
        let answer = storedItems[userAdId].answerList[i];
        resultArray.push({"role": "user", "content": question});
        resultArray.push({"role": "assistant", "content": answer});
      }
      return resultArray;
    }
    else {
      return [];
    }
  }
  catch (err){
    console.log(`Read rejected. ${err}`);
    return [];
  }
}