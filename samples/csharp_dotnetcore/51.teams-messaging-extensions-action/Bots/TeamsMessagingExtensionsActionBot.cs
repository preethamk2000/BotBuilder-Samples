// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//extern alias BetaLib;

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.BotBuilderSamples.Helpers;
using Microsoft.BotBuilderSamples.Models;
using TeamsMessagingExtensionsAction.Model;
using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Identity.Client;
using System.Text;
//using Beta = BetaLib.Microsoft.Graph.Beta;
//using MSGraphBeta = Microsoft.Graph.Beta;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class TeamsMessagingExtensionsActionBot : TeamsActivityHandler
    {
        public readonly string baseUrl;
        public readonly string clientID;
        public readonly string clientSecret;
        public readonly string graphAPIToken;

        public TeamsMessagingExtensionsActionBot(IConfiguration configuration) : base()
        {
            this.baseUrl = configuration["BaseUrl"];
            this.clientID = configuration["MicrosoftAppId"];
            this.clientSecret = configuration["MicrosoftAppPassword"];
            this.graphAPIToken = configuration["GraphAPIToken"];
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var infoText = "Please specify in this format: scheduleMessageLater AliasName [Message] [dd/MM HH:mm:ss]";

            if (turnContext.Activity.Text != null)
            {
                var text = turnContext.Activity.Text.Trim();

                if(text.Equals("schedulemessagelater --help"))
                    await turnContext.SendActivityAsync(MessageFactory.Text(infoText, infoText), cancellationToken);

                else if (text.Contains("scheduleMessageLater"))
                    await ScheduleMessageLater(turnContext, cancellationToken, text);
            }

        }

        private async Task ScheduleMessageLater(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken, string text)
        {
            string alias = text.Split(" ")[1].ToLower();
            string message = text.Split('[')[1].Split(']')[0];
            string dateTimeString = text.Split('[')[2].Split(']')[0];

            //var enteredString = "14/10 21:00:00";
            DateTime myDate = DateTime.ParseExact(dateTimeString, "dd/MM HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
            myDate = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(myDate, "India Standard Time", "UTC");
            int secondsToSend = ((int)(myDate - DateTime.UtcNow).TotalSeconds);

            //var replyText = $"**Message:** {message} {Environment.NewLine} **Scheduled at:** {dateTimeString} {Environment.NewLine} **To:** {alias}";
            //await turnContext.SendActivityAsync(MessageFactory.Text(replyText, replyText), cancellationToken);
            var responseCard = GetCardForResponse(message, alias, dateTimeString);
            await turnContext.SendActivityAsync(responseCard, cancellationToken);

            try
            {
                var chatID = GetChatIDFromAliasManual(alias);
                Thread.Sleep(secondsToSend * 1000);
                dynamic response = SendMessageWithChatID(message, chatID);
                Console.WriteLine(response.body.content);
                await turnContext.SendActivityAsync(MessageFactory.Text($"Message to {alias} has been sent!"), cancellationToken);
            }
            catch
            {
                await turnContext.SendActivityAsync(MessageFactory.Text($"Failed to send the message to {alias}."), cancellationToken);
            }

            //await turnContext.SendActivityAsync(MessageFactory.Text(chatID), cancellationToken);


        }

        private object GetObjectFromGraphAPI(string url)
        {
            dynamic response = null;
            try
            {
                var webRequest = System.Net.WebRequest.Create(url);
                if (webRequest != null)
                {
                    webRequest.Method = "GET";
                    webRequest.Timeout = 15000;
                    webRequest.ContentType = "application/json";
                    webRequest.Headers.Add("Authorization", $"Bearer {this.graphAPIToken}");

                    using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                    {
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                        {
                            var jsonResponse = sr.ReadToEnd();
                            response = JsonConvert.DeserializeObject(jsonResponse);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return response;
        }

        private string GetChatIDFromAliasManual(string alias)
        {
            string WEBSERVICE_URL = "https://graph.microsoft.com/beta/me/chats/";
            dynamic response = GetObjectFromGraphAPI(WEBSERVICE_URL);
            string chatID = null;
            //Console.WriteLine(String.Format("Response: {0}", response));
            //Console.WriteLine(response["@odata.count"]);
            foreach (var chat in response.value)
            {
                if (chat.chatType == "oneOnOne")
                {
                    string url = $"https://graph.microsoft.com/beta/chats/{chat.id}/members";
                    dynamic chatInfo = GetObjectFromGraphAPI(url);
                    if (Int32.Parse(chatInfo["@odata.count"].ToString()) == 2 && (chatInfo.value[1].email.ToString().Split("@")[0].ToLower() == alias || chatInfo.value[0].email.ToString().Split("@")[0].ToLower() == alias) )
                    {
                        //Console.WriteLine(chat.id);
                        chatID = chat.id.ToString();
                        break;
                    }
                }
            }

            return chatID;
        }

        private object SendMessageWithChatID(string message, string chatID)
        {
            dynamic response = null;
            try
            {
                var webRequest = System.Net.WebRequest.Create($"https://graph.microsoft.com/v1.0/chats/{chatID}/messages");
                if (webRequest != null)
                {
                    webRequest.Method = "POST";
                    webRequest.Timeout = 15000;
                    webRequest.ContentType = "application/json";
                    webRequest.Headers.Add("Authorization", $"Bearer {this.graphAPIToken}");

                    string stringData = " {\"body\": {\"content\": \" " + message + " \"}  } ";
                    var data = Encoding.Default.GetBytes(stringData); // note: choose appropriate encoding
                    webRequest.ContentLength = data.Length;

                    var newStream = webRequest.GetRequestStream(); // get a ref to the request body so it can be modified
                    newStream.Write(data, 0, data.Length);
                    newStream.Close();

                    using (System.IO.Stream s = webRequest.GetResponse().GetResponseStream())
                    {
                        using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                        {
                            var jsonResponse = sr.ReadToEnd();
                            response = JsonConvert.DeserializeObject(jsonResponse);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return response;
        }

        protected IMessageActivity GetCardForResponse(String message, string alias, string dateTimeString)
        {

            var card = new HeroCard();

            card.Title = "Scheduled Message Details";
            card.Subtitle = $"To: {alias}, At: {dateTimeString}";
            card.Text = $"Message: {message}";

            var activity = MessageFactory.Attachment(card.ToAttachment());
            return activity;

        }

        private MessagingExtensionActionResponse CreateCardCommand(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // The user has chosen to create a card by choosing the 'Create Card' context menu command.
            var createCardData = ((JObject)action.Data).ToObject<CardResponse>();

            var card = new HeroCard
            {
                Title = createCardData.Title,
                Subtitle = createCardData.Subtitle,
                Text = createCardData.Text,
            };

            var attachments = new List<MessagingExtensionAttachment>();
            attachments.Add(new MessagingExtensionAttachment
            {
                Content = card,
                ContentType = HeroCard.ContentType,
                Preview = card.ToAttachment(),
            });

            return new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    AttachmentLayout = "list",
                    Type = "result",
                    Attachments = attachments,
                },
            };
        }

        private MessagingExtensionActionResponse ShareMessageCommand(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // The user has chosen to share a message by choosing the 'Share Message' context menu command.
            var heroCard = new HeroCard
            {
                Title = $"{action.MessagePayload.From?.User?.DisplayName} orignally sent this message:",
                Text = action.MessagePayload.Body.Content,
            };

            if (action.MessagePayload.Attachments != null && action.MessagePayload.Attachments.Count > 0)
            {
                // This sample does not add the MessagePayload Attachments.  This is left as an
                // exercise for the user.
                heroCard.Subtitle = $"({action.MessagePayload.Attachments.Count} Attachments not included)";
            }

            // This Messaging Extension example allows the user to check a box to include an image with the
            // shared message.  This demonstrates sending custom parameters along with the message payload.
            var includeImage = ((JObject)action.Data)["includeImage"]?.ToString();
            if (string.Equals(includeImage, bool.TrueString, StringComparison.OrdinalIgnoreCase))
            {
                heroCard.Images = new List<CardImage>
                {
                    new CardImage { Url = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU" },
                };
            }

            return new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = new List<MessagingExtensionAttachment>()
                    {
                        new MessagingExtensionAttachment
                        {
                            Content = heroCard,
                            ContentType = HeroCard.ContentType,
                            Preview = heroCard.ToAttachment(),
                        },
                    },
                },
            };
        }

        private MessagingExtensionActionResponse WebViewResponse(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // The user has chosen to create a card by choosing the 'Web View' context menu command.
            CustomFormResponse cardData = JsonConvert.DeserializeObject<CustomFormResponse>(action.Data.ToString());
            var imgUrl = baseUrl + "/MSFT_logo.jpg";
            var card = new ThumbnailCard
            {
                Title = "ID: " + cardData.EmpId,
                Subtitle = "Name: " + cardData.EmpName,
                Text = "E-Mail: " + cardData.EmpEmail,
                Images = new List<CardImage> { new CardImage { Url = imgUrl } },
            };

            var attachments = new List<MessagingExtensionAttachment>();
            attachments.Add(new MessagingExtensionAttachment
            {
                Content = card,
                ContentType = ThumbnailCard.ContentType,
                Preview = card.ToAttachment(),
            });

            return new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    AttachmentLayout = "list",
                    Type = "result",
                    Attachments = attachments,
                },
            };
        }

        private MessagingExtensionActionResponse CreateAdaptiveCardResponse(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            var createCardResponse = ((JObject)action.Data).ToObject<CardResponse>();
            var attachments = CardHelper.CreateAdaptiveCardAttachment(action, createCardResponse);

            return new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    AttachmentLayout = "list",
                    Type = "result",
                    Attachments = attachments,
                },
            };
        }

        private MessagingExtensionActionResponse DateDayInfo(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            var response = new MessagingExtensionActionResponse()
            {
                Task = new TaskModuleContinueResponse()
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Height = 175,
                        Width = 300,
                        Title = "Task Module Razor View",
                        Url = baseUrl + "/Home/RazorView",
                    },
                },
            };
            return response;
        }

        private MessagingExtensionActionResponse TaskModuleHTMLPage(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            var response = new MessagingExtensionActionResponse()
            {
                Task = new TaskModuleContinueResponse()
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Height = 200,
                        Width = 400,
                        Title = "Task Module HTML Page",
                        Url = baseUrl + "/htmlpage.html",
                    },
                },
            };
            return response;
        }

        private MessagingExtensionActionResponse EmpDetails(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            var response = new MessagingExtensionActionResponse()
            {
                Task = new TaskModuleContinueResponse()
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Height = 300,
                        Width = 450,
                        Title = "Task Module WebView",
                        Url = baseUrl + "/Home/CustomForm",
                    },
                },
            };
            return response;
        }
    }
}
