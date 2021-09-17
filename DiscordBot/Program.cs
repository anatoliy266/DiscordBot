using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

using DSharpPlus;
using DSharpPlus.SlashCommands;
using DSharpPlus.Entities;
using DSharpPlus.Lavalink;

using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using Google.Apis.Sheets.v4.Data;
using static Google.Apis.Sheets.v4.SpreadsheetsResource;
using System.Linq;

namespace ExampleGatewayBot
{
    class Program
    {
        public static DiscordClient Discord;
        static void Main(string[] args)
        {
            
            // ... create a new DiscordClient for the bot ...
            Discord = new DiscordClient(new DiscordConfiguration
            {
                Token = "xxxxxx",
                TokenType = TokenType.Bot,
                ShardCount = 1,
                Intents = DiscordIntents.AllUnprivileged,
                MinimumLogLevel = Microsoft.Extensions.Logging.LogLevel.Debug
            });

            Discord.MessageCreated += Discord_MessageCreated;

            var slash = Discord.UseSlashCommands();
            slash.RegisterCommands<SlashCommands>();
            // ... connect to discord ...
            Discord.ConnectAsync().GetAwaiter().GetResult();
            // ... and prevent this from stopping.
            Task.Delay(-1).GetAwaiter().GetResult();
        }

        private static Task Discord_MessageCreated(DiscordClient sender, DSharpPlus.EventArgs.MessageCreateEventArgs e)
        {
            if (e.Author.IsCurrent)
            {
                if (e.Message.Content.Contains("Приветствую, Салага"))
                {
                    sender.SendMessageAsync(e.Channel, "Экзамен отменяется потомучто ты гомосек, азазаза :)");
                    return Task.CompletedTask;
                } 
                else
                {
                    return Task.CompletedTask;
                }
            } else
            {
                return Task.CompletedTask;
            }
        }

        private static Task Discord_ApplicationCommandUpdated(DiscordClient sender, DSharpPlus.EventArgs.ApplicationCommandEventArgs e)
        {
            Discord.Logger.LogInformation($"Shard {sender.ShardId} sent application command updated: {e.Command.Name}: {e.Command.Id} for {e.Command.ApplicationId}");
            return Task.CompletedTask;
        }
        private static Task Discord_ApplicationCommandDeleted(DiscordClient sender, DSharpPlus.EventArgs.ApplicationCommandEventArgs e)
        {
            Discord.Logger.LogInformation($"Shard {sender.ShardId} sent application command deleted: {e.Command.Name}: {e.Command.Id} for {e.Command.ApplicationId}");
            return Task.CompletedTask;
        }
        private static Task Discord_ApplicationCommandCreated(DiscordClient sender, DSharpPlus.EventArgs.ApplicationCommandEventArgs e)
        {
            Discord.Logger.LogInformation($"Shard {sender.ShardId} sent application command created: {e.Command.Name}: {e.Command.Id} for {e.Command.ApplicationId}");
            return Task.CompletedTask;
        }
    }

    public class SlashCommands : ApplicationCommandModule
    {
        [SlashCommand("Play", "Играть музыку")]
        public async Task Play(InteractionContext ctx, [Option("link", "ссылка на музыку")] DiscordUser user)
        {
            throw new NotImplementedException();
        }

        [SlashCommand("Start", "Начать экзамен")]
        public async Task Start(InteractionContext ctx, [Option("User", "Экзаменующийся")] DiscordUser user)
        {
            try
            {
                var api = new SheetApi("xxxxxxxxxx");

                var data = api.ReadData("Ответы!A:L");
                
                if (data.Where(x => (string)x[1] == user.Username).Count() == 0)
                {
                    await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().WithContent($"Приветствую, Салага \r\n Это твоя первая попытка провалить экзамен на должность рядового \r\n Условия простые - 10 вопросов и 15 минут \r\n Приступить к выполнению!"));
                } else if (data.Where(x => (string)x[1] == user.Username).Count() == 1)
                {
                    await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().WithContent($"Приветствую, Салага \r\n Это твоя вторая и последняя попытка, салага \r\n Условия те же - 15 минут и 10 вопросов \r\n Если не справишься - будешь уволен"));
                } else
                {
                    await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().WithContent($"Все попытки пройти экзамен исчерпаны, пошел нахуй"));
                }
            } catch (Exception e)
            {
                var m = e.Message;
                await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().WithContent($"Я заболел, пошел нахуй"));
            }
        }

        [SlashCommand("Invite", "Отправить пользователя на экзамен")]
        public async Task Invite(InteractionContext ctx, [Option("User", "Экзаменующийся")] DiscordUser user)
        {
            try
            {
                var api = new SheetApi("xxxxxxxxx");

                var data = api.ReadData("Ответы!A:L");

                if (data.Where(x => (string)x[1] == user.Username).Count() < 2)
                {
                    var users = ctx.Client.Guilds.Values.Where(x => x.Id == ctx.Guild.Id).FirstOrDefault().Members;
                    var ch = ctx.Channel;
                    var guid = Guid.NewGuid();
                    api.WriteData("Ответы", data.Count + 1, "A", guid.ToString());
                    api.WriteData("Ответы", data.Count + 1, "B", user.Username);
                    api.WriteData("Ответы", data.Count + 1, "C", ctx.Guild.Members.Values.ToList().Where(x => x.Id == user.Id).FirstOrDefault()?.DisplayName);
                    await ctx.Guild.Members.Where(x => x.Value.Id == user.Id).FirstOrDefault().Value.SendMessageAsync("Быстро написал /start или всю службу будешь траву красить и очки чистить!!!!!");
                    await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().WithContent($"Очередное мясо в мою коллекцию, {ctx.Guild.Members.Where(x => x.Value.Id == user.Id).FirstOrDefault().Value.DisplayName} за мной в отдельный канал!!!!"));
                }
                else
                {
                    await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().WithContent($"Все попытки пройти экзамен исчерпаны, пошел нахуй"));
                }
            } catch (Exception e)
            {
                var m = e.Message;
                await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().WithContent($"{ctx.User}, Я заболел, пошел нахуй"));
            }
        }
    }


    public class SheetApi
    {
        public SheetsService Service;

        public string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public string ApplicationName = "xxxxxxxx";
        public string SheetId { get; set; }
        public SheetApi(string sheetId)
        {
            UserCredential credential;
            SheetId = sheetId;
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            Service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }


        public IList<IList<Object>> ReadData(string sheetName, int recId)
        {
            String range = $"{sheetName}!A{recId}";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    Service.Spreadsheets.Values.Get(SheetId, range);
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            return values;
        }

        public IList<IList<Object>> ReadData(string range)
        {
            var request = Service.Spreadsheets.Values.Get(SheetId, range);
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            return values;
        }

        public void WriteData(string sheetName, int row, string cell, string content)
        {
            var range = $"{sheetName}!{cell}{row}";
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

            var oblist = new List<object>() { content };
            valueRange.Values = new List<IList<object>> { oblist };

            var update = Service.Spreadsheets.Values.Update(valueRange, SheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            update.Execute();
        }
    }
}