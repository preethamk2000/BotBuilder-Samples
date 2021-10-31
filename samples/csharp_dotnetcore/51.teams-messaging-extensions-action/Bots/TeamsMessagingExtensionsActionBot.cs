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

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
           switch (action.CommandId)
           {
               case "scheduleMessage":
                   return ScheduleMessageLater(turnContext, action, cancellationToken);
               //case "shareMessage":
               //    return ShareMessageCommand(turnContext, action);
               //case "webView":
               //    return WebViewResponse(turnContext, action);
               //case "createAdaptiveCard":
               //    return CreateAdaptiveCardResponse(turnContext, action);
               //case "razorView":
               //    return RazorViewResponse(turnContext, action);
           }
           return new MessagingExtensionActionResponse();
        }

        private async Task SendScheduledMessageResponse(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken, string replyText)
        {
            //var activity = GetCardForNewReminder(outputString);
            // Echo back what the user said
            await turnContext.SendActivityAsync(MessageFactory.Text(replyText), cancellationToken);
        }

        private MessagingExtensionActionResponse ScheduleMessageLater(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
           // The user has chosen to create a card by choosing the 'Create Card' context menu command.
           var scheduleMessageData = ((JObject)action.Data).ToObject<ScheduleMessageResponse>();

           var chatId = GetChatIDFromAlias(scheduleMessageData.Recipient);
           var replyText = $"Message: {scheduleMessageData.Message} | Scheduled at: <> | To: {scheduleMessageData.Recipient}";

           var response = GetCardForResponse(replyText);

           SendScheduledMessageResponse(turnContext, cancellationToken, replyText);


           var card = new HeroCard
           {
              Title = "Scheduled Message Details",
              Subtitle = $"To: {scheduleMessageData.Recipient}",
              Text = scheduleMessageData.Message,
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

           return new MessagingExtensionActionResponse();
        }

        private GraphServiceClient GetGraphClient()
        {
            var scopes = new[] { "User.Read", "Chat.ReadBasic", "ChatMember.Read", "ChatMessage.Send" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "common";

            //// Values from app registration
            var clientId = this.clientID;
            var clientSecret = "YOUR_CLIENT_SECRET";

            var options = new TokenCredentialOptions
            {
               AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // This is the incoming token to exchange using on-behalf-of flow
            var oboToken = "JWT_TOKEN_TO_EXCHANGE";

            var cca = ConfidentialClientApplicationBuilder
               .Create(this.clientID)
               .WithTenantId(tenantId)
               .WithClientSecret(this.clientSecret)
               .Build();

            // DelegateAuthenticationProvider is a simple auth provider implementation
            // that allows you to define an async function to retrieve a token
            // Alternatively, you can create a class that implements IAuthenticationProvider
            // for more complex scenarios
            var authProvider = new DelegateAuthenticationProvider(async (request) => {
               // Use Microsoft.Identity.Client to retrieve token
               var assertion = new UserAssertion(oboToken);
               var result = await cca.AcquireTokenOnBehalfOf(scopes, assertion).ExecuteAsync();

               request.Headers.Authorization =
                   new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", result.AccessToken);
            });

            var scopes = new[] { "https://graph.microsoft.com/Chat.ReadBasic.All", "https://graph.microsoft.com/Chat.ReadWrite.All", "https://graph.microsoft.com/ChatMember.Read.All", "https://graph.microsoft.com/User.Read" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "common";


            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, this.clientID, this.clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            //var graphClient = new GraphServiceClient(authProvider);
            return graphClient;
        }

        private async Task<string> GetChatIDFromAlias(string recipient)
        {
            var client = GetGraphClient();
            var chats = await client.Me.Chats.Request().GetAsync();

            //foreach (var chat in chats)
            //{
            //    var chatID = chat.Id;

            //    var members = await client.Chats[chatID].Members
            //                        .Request()
            //                        .GetAsync();


            //}
            return chats.ToString();
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
