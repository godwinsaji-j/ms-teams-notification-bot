var notificationTemplate = require("./adaptiveCards/notification-default.json");
var notificationAlertTemplate = require("./adaptiveCards/notification-card-bg.json");
var notificationReportTemplate = require("./adaptiveCards/notification-report.json");
var notificationReportIssueTemplate = require("./adaptiveCards/notification-card-report-issue.json");
const { bot } = require("./internal/initialize");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const restify = require("restify");
const axios = require('axios').default;

const ngrokURL = 'https://ee95-2401-4900-6050-4f-99da-b0dd-2cd-e6b.in.ngrok.io';

// Create HTTP server.
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
server.post(
  "/api/notification",
  restify.plugins.queryParser(),
  restify.plugins.bodyParser(), // Add more parsers if needed
  async (req, res) => {
    var splunkEventData = {type:"Normal Notification",sent:[],ignored:[],botNotInstalled:[]};
    console.log("Inside Bot Line 1");
    var notificationRecipients=[];
    var graphApiUserList = require('./graphApi');
    var useraadObjectIds = Object.keys(graphApiUserList);
    var data = JSON.parse(req.body)
    var activityId = data.activityId;
    var title = data.activityName;
    var description = data.activityDescription;
    var followedBy = data.trackedBy;
    var trackedForEmailIds =  data.trackedForEmailIds;
    console.log("Assignments Completed in Bot");
    if(data.notificationTemplate){
      notificationTemplate = data.notificationTemplate
    }
    for (const target of await bot.notification.installations()) {
      var currentUseraadObjectId = target.conversationReference.user.aadObjectId;
      console.log("Inside Bot for loop, currentUserObjectId = ",currentUseraadObjectId);
      if(graphApiUserList[currentUseraadObjectId]){
        console.log("Valid Match Found for Tenant Id in out Otganization");
        var currentUserEmailId = graphApiUserList[currentUseraadObjectId].mail;
        notificationRecipients.push(currentUserEmailId);
        if(trackedForEmailIds.includes(currentUserEmailId)){
          splunkEventData.sent.push(currentUserEmailId);
          console.log("Sending Adaptive Card");
          await target.sendAdaptiveCard(
            AdaptiveCards.declare(notificationTemplate).render({
              title: "New Event Occurred! - "+title,
              followedBy: `From: ${followedBy}`,
              addressingUser: `Hi ${graphApiUserList[currentUseraadObjectId].displayName}, `,
              description: `${description}`,
              completedNotificationUrl: `${ngrokURL}/activity/mark-task-completed/${activityId}/${currentUserEmailId}`,
              inProgressNotificationUrl:`${ngrokURL}/activity/mark-task-progress/${activityId}/${currentUserEmailId}`,
              reportIssueUrl: `${ngrokURL}/activity/report-task-issue/${activityId}/${currentUserEmailId}`,
              adaptiveCardBGImg: `https://i.postimg.cc/NjxK4V4D/blue-adaptive-card-bg.jpg`
            })
          );
          trackedForEmailIds.splice(trackedForEmailIds.indexOf(currentUserEmailId),1)
        }else{
          splunkEventData.ignored.push(currentUserEmailId);
          console.log("Ignoring "+currentUserEmailId+", Since notification is not targetted");  
        }
        
      }else{
        console.log("Tenant Id Belongs to our Organization");
      }
      
    }
    //Splunk Logging
    splunkEventData.botNotInstalled = trackedForEmailIds;
    console.log("Posting data log to Splunk",splunkEventData);
    await axios.post(ngrokURL+'/activity/log-splunk-event', splunkEventData)
    .then(function (response) {
      console.log("Splunk Api Response = ",response);
    })
    .catch(function (error) {
      console.log(error);
    });


    /****** To distinguish different target types ******/
    /** "Channel" means this bot is installed to a Team (default to notify General channel)
    if (target.type === NotificationTargetType.Channel) {
      // Directly notify the Team (to the default General channel)
      await target.sendAdaptiveCard(...);

      // List all channels in the Team then notify each channel
      const channels = await target.channels();
      for (const channel of channels) {
        await channel.sendAdaptiveCard(...);
      }

      // List all members in the Team then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Group" means this bot is installed to a Group Chat
    if (target.type === NotificationTargetType.Group) {
      // Directly notify the Group Chat
      await target.sendAdaptiveCard(...);

      // List all members in the Group Chat then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Person" means this bot is installed as a Personal app
    if (target.type === NotificationTargetType.Person) {
      // Directly notify the individual person
      await target.sendAdaptiveCard(...);
    }
    **/
    //res.body={notificationRecipients:notificationRecipients};
    res.json();
  }
);

server.post(
  "/api/AlertNotification",
  restify.plugins.queryParser(),
  restify.plugins.bodyParser(), // Add more parsers if needed
  async (req, res) => {
    var splunkEventData = {type:"Normal Notification",sent:[],ignored:[],botNotInstalled:[]};
    console.log("Inside Bot Line 1");
    var notificationRecipients=[];
    var graphApiUserList = require('./graphApi');
    var useraadObjectIds = Object.keys(graphApiUserList);
    var data = JSON.parse(req.body)
    var activityId = data.activityId;
    var title = data.activityName;
    var description = data.activityDescription;
    var followedBy = data.trackedBy;
    var trackedForEmailIds =  data.trackedForEmailIds;
    var completedDL = data.completedDL;
    //Alert Type and Ignoring Completed Users
    var alertType = data.alertType;
    alertType = (alertType)?alertType:"";
    splunkEventData.type = alertType
    var adaptiveCardHeaderIcon = "";
    var adaptiveCardBGImg = "";
    var ReminderType = "";
    if(alertType == "warning"){
      ReminderType = "Reminder";
      adaptiveCardHeaderIcon = "https://i.postimg.cc/P5yF4nPN/warning.png"
      adaptiveCardBGImg = `https://i.postimg.cc/VkxTwNPD/yellow-warning-adaptive-card-bg.jpg`
    }else{
      //Urgent By Default
      ReminderType = "Final Reminder";
      adaptiveCardHeaderIcon = "https://i.postimg.cc/KvRX4zzQ/bell.png";
      adaptiveCardBGImg = `https://i.postimg.cc/YS9CxSj2/red-urgent-adaptive-card-bg.jpg`
    }
    console.log("Card BG Image",adaptiveCardBGImg);

    for (const target of await bot.notification.installations()) {
      var currentUseraadObjectId = target.conversationReference.user.aadObjectId;
      console.log("Inside Bot for loop, currentUserObjectId = ",currentUseraadObjectId);
      if(graphApiUserList[currentUseraadObjectId]){
        console.log("Valid Match Found for Tenant Id in out Otganization");
        var currentUserEmailId = graphApiUserList[currentUseraadObjectId].mail;
        notificationRecipients.push(currentUserEmailId);
        if(trackedForEmailIds.includes(currentUserEmailId) && !completedDL.includes(currentUserEmailId)){
          splunkEventData.sent.push(currentUserEmailId);
          console.log("Sending Adaptive Card");
          await target.sendAdaptiveCard(
            AdaptiveCards.declare(notificationAlertTemplate).render({
              title: ReminderType+"! - "+title,
              adaptiveCardHeaderIcon:`${adaptiveCardHeaderIcon}`,
              followedBy: `From: ${followedBy}`,
              addressingUser: `Hi ${graphApiUserList[currentUseraadObjectId].displayName}, `,
              description: `${description}`,
              completedNotificationUrl: `${ngrokURL}/activity/mark-task-completed/${activityId}/${currentUserEmailId}`,
              inProgressNotificationUrl:`${ngrokURL}/activity/mark-task-progress/${activityId}/${currentUserEmailId}`,
              reportIssueUrl: `${ngrokURL}/activity/report-task-issue/${activityId}/${currentUserEmailId}`,
              adaptiveCardBGImg:adaptiveCardBGImg
            })
          );
          trackedForEmailIds.splice(trackedForEmailIds.indexOf(currentUserEmailId),1)
        }else if(completedDL.includes(currentUserEmailId)){
          splunkEventData.ignored.push(`(completed)${currentUserEmailId}`);
        }
        else{
          splunkEventData.ignored.push(`(not in followUp)${currentUserEmailId}`);
          console.log("Ignoring "+currentUserEmailId+", Since notification is not targetted");  
        }
        
      }else{
        console.log("Tenant Id Belongs to our Organization");
      }
      
    }
    //Splunk Logging
    splunkEventData.botNotInstalled = trackedForEmailIds;
    console.log("Posting data log to Splunk",splunkEventData);
    await axios.post(ngrokURL+'/activity/log-splunk-event', splunkEventData)
    .then(function (response) {
      console.log("Splunk Api Response = ",response);
    })
    .catch(function (error) {
      console.log(error);
    });


    /****** To distinguish different target types ******/
    /** "Channel" means this bot is installed to a Team (default to notify General channel)
    if (target.type === NotificationTargetType.Channel) {
      // Directly notify the Team (to the default General channel)
      await target.sendAdaptiveCard(...);

      // List all channels in the Team then notify each channel
      const channels = await target.channels();
      for (const channel of channels) {
        await channel.sendAdaptiveCard(...);
      }

      // List all members in the Team then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Group" means this bot is installed to a Group Chat
    if (target.type === NotificationTargetType.Group) {
      // Directly notify the Group Chat
      await target.sendAdaptiveCard(...);

      // List all members in the Group Chat then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Person" means this bot is installed as a Personal app
    if (target.type === NotificationTargetType.Person) {
      // Directly notify the individual person
      await target.sendAdaptiveCard(...);
    }
    **/
    //res.body={notificationRecipients:notificationRecipients};
    res.json();
  }
);

server.post(
  "/api/ReportImageNotification",
  restify.plugins.queryParser(),
  restify.plugins.bodyParser(), // Add more parsers if needed
  async (req, res) => {
    var splunkEventData = {type:"Normal Notification",sent:[],ignored:[],botNotInstalled:[]};
    console.log("Inside Bot Line 1");
    var notificationRecipients=[];
    var graphApiUserList = require('./graphApi');
    var useraadObjectIds = Object.keys(graphApiUserList);
    var data = JSON.parse(req.body)
    var activityId = data.activityId;
    var title = data.activityName;
    var description = data.activityDescription;
    var followedBy = data.trackedBy;
    var trackedForEmailIds =  data.trackedForEmailIds;
    var completedDL = data.completedDL;
    var leadsDL = data.leadsDL;
    //Alert Type and Ignoring Completed Users
    var alertType = data.alertType;
    alertType = (alertType)?alertType:"";
    splunkEventData.type = alertType
    var adaptiveCardReportChartImg = data.reportDataImage;
    console.log("Assignments Completed in Bot");
    
    for (const target of await bot.notification.installations()) {
      var currentUseraadObjectId = target.conversationReference.user.aadObjectId;
      console.log("Inside Bot for loop, currentUserObjectId = ",currentUseraadObjectId);
      if(graphApiUserList[currentUseraadObjectId]){
        console.log("Valid Match Found for Tenant Id in out Otganization");
        var currentUserEmailId = graphApiUserList[currentUseraadObjectId].mail;
        notificationRecipients.push(currentUserEmailId);
        "Send Report Only to Leads"
        if(leadsDL.includes(currentUserEmailId)){
          splunkEventData.sent.push(currentUserEmailId);
          console.log("Sending Adaptive Card");
          await target.sendAdaptiveCard(
            AdaptiveCards.declare(notificationReportTemplate).render({
              title: "Report for - "+title+"!",
              followedBy: `From: ${followedBy}`,
              addressingUser: `Hi ${graphApiUserList[currentUseraadObjectId].displayName}, `,
              description: `${description}`,
              completedNotificationUrl: `${ngrokURL}/activity/mark-task-completed/${activityId}/${currentUserEmailId}`,
              inProgressNotificationUrl:`${ngrokURL}/activity/mark-task-progress/${activityId}/${currentUserEmailId}`,
              reportIssueUrl: `${ngrokURL}/activity/report-task-issue/${activityId}/${currentUserEmailId}`,
              adaptiveCardReportChartImg:adaptiveCardReportChartImg,
              adaptiveCardBGImg:`https://i.postimg.cc/NjxK4V4D/blue-adaptive-card-bg.jpg`
            })
          );
          trackedForEmailIds.splice(trackedForEmailIds.indexOf(currentUserEmailId),1)
        }else if(completedDL.includes(currentUserEmailId)){
          splunkEventData.ignored.push(`(completed)${currentUserEmailId}`);
        }
        else{
          splunkEventData.ignored.push(`(not in followUp)${currentUserEmailId}`);
          console.log("Ignoring "+currentUserEmailId+", Since notification is not targetted");  
        }
        
      }else{
        console.log("Tenant Id Belongs to our Organization");
      }
      
    }
    //Splunk Logging
    splunkEventData.botNotInstalled = trackedForEmailIds;
    console.log("Posting data log to Splunk",splunkEventData);
    await axios.post(ngrokURL+'/activity/log-splunk-event', splunkEventData)
    .then(function (response) {
      console.log("Splunk Api Response = ",response);
    })
    .catch(function (error) {
      console.log(error);
    });


    /****** To distinguish different target types ******/
    /** "Channel" means this bot is installed to a Team (default to notify General channel)
    if (target.type === NotificationTargetType.Channel) {
      // Directly notify the Team (to the default General channel)
      await target.sendAdaptiveCard(...);

      // List all channels in the Team then notify each channel
      const channels = await target.channels();
      for (const channel of channels) {
        await channel.sendAdaptiveCard(...);
      }

      // List all members in the Team then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Group" means this bot is installed to a Group Chat
    if (target.type === NotificationTargetType.Group) {
      // Directly notify the Group Chat
      await target.sendAdaptiveCard(...);

      // List all members in the Group Chat then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Person" means this bot is installed as a Personal app
    if (target.type === NotificationTargetType.Person) {
      // Directly notify the individual person
      await target.sendAdaptiveCard(...);
    }
    **/
    //res.body={notificationRecipients:notificationRecipients};
    res.json();
  }
);

server.post(
  "/api/ReportBugNotification",
  restify.plugins.queryParser(),
  restify.plugins.bodyParser(), // Add more parsers if needed
  async (req, res) => {
    var splunkEventData = {type:"Report Bug Notification To Tracker",sent:[],ignored:[],botNotInstalled:[]};
    console.log("Inside Bot Line 1");
    var notificationRecipients=[];
    var graphApiUserList = require('./graphApi');
    var useraadObjectIds = Object.keys(graphApiUserList);
    var data = JSON.parse(req.body)
    var activityId = data.activityId;
    var title = data.activityName;
    var description = data.issueDescription;
    var issueReportedBy = data.issueReportedBy;
    var reportIssuesDL =  data.reportIssuesDL;
    //Alert Type and Ignoring Completed Users
      var adaptiveCardHeaderIcon = "https://i.postimg.cc/tgfkSqyR/bug.png"
      var adaptiveCardBGImg = `https://i.postimg.cc/HkjLwxrY/bug-adaptive-card-bg.jpg`
      console.log("Card BG Image",adaptiveCardBGImg);

    for (const target of await bot.notification.installations()) {
      var currentUseraadObjectId = target.conversationReference.user.aadObjectId;
      console.log("Inside Bot for loop, currentUserObjectId = ",currentUseraadObjectId);
      if(graphApiUserList[currentUseraadObjectId]){
        console.log("Valid Match Found for Tenant Id in out Otganization");
        var currentUserEmailId = graphApiUserList[currentUseraadObjectId].mail;
        notificationRecipients.push(currentUserEmailId);
        if(reportIssuesDL.includes(currentUserEmailId)){
          splunkEventData.sent.push(currentUserEmailId);
          console.log("Sending Adaptive Card");
          await target.sendAdaptiveCard(
            AdaptiveCards.declare(notificationReportIssueTemplate).render({
              title: "Issue Reported In! - "+title,
              adaptiveCardHeaderIcon:`${adaptiveCardHeaderIcon}`,
              issueReportedBy: `From: ${issueReportedBy}`,
              addressingUser: `Hi ${graphApiUserList[currentUseraadObjectId].displayName}, `,
              genricIssueReportingText: `Please find below the Issue/Bug details for this Activity${description}`,
              description:`${description}`,
              adaptiveCardBGImg:adaptiveCardBGImg
            })
          );
          reportIssuesDL.splice(reportIssuesDL.indexOf(currentUserEmailId),1)
        }
      }else{
        console.log("Tenant Id Belongs to our Organization");
      }
      
    }
    //Splunk Logging
    splunkEventData.botNotInstalled = reportIssuesDL;
    console.log("Posting data log to Splunk",splunkEventData);
    await axios.post(ngrokURL+'/activity/log-splunk-event', splunkEventData)
    .then(function (response) {
      console.log("Splunk Api Response = ",response);
    })
    .catch(function (error) {
      console.log(error);
    });


    /****** To distinguish different target types ******/
    /** "Channel" means this bot is installed to a Team (default to notify General channel)
    if (target.type === NotificationTargetType.Channel) {
      // Directly notify the Team (to the default General channel)
      await target.sendAdaptiveCard(...);

      // List all channels in the Team then notify each channel
      const channels = await target.channels();
      for (const channel of channels) {
        await channel.sendAdaptiveCard(...);
      }

      // List all members in the Team then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Group" means this bot is installed to a Group Chat
    if (target.type === NotificationTargetType.Group) {
      // Directly notify the Group Chat
      await target.sendAdaptiveCard(...);

      // List all members in the Group Chat then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Person" means this bot is installed as a Personal app
    if (target.type === NotificationTargetType.Person) {
      // Directly notify the individual person
      await target.sendAdaptiveCard(...);
    }
    **/
    //res.body={notificationRecipients:notificationRecipients};
    res.json();
  }
);

// Bot Framework message handler.
server.post("/api/messages", async (req, res) => {
  await bot.requestHandler(req, res);
});
