using ClosedXML.Excel;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

namespace MyGridBot
{
    internal class TG
    {
        public static int Sorting = 50;
        public static string Token = "";
        public static string Notify = "";
        public static string Bashorg = "";
        public static string Buttons = "";
        public static string Report = "📊 Подготовка отчета";
        public static TelegramBotClient Client = new(Token);
        public static Chat Chat = new();
        private static readonly string PathTG = @"..\\..\\..\\..\\Telegram.xlsx";
        private static ITelegramBotClient _botClient;
        private static ReceiverOptions _receiverOptions;
        private static readonly HttpClient HttpClient = new();

        public static async Task SendMessageAsync(string message)
        {
            if (!string.IsNullOrEmpty(Token))
            {
                int retryCount = 0;
                while (retryCount < 3)
                {
                    try
                    {
                        await Client.SendTextMessageAsync(Chat.Id, message);
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при отправке сообщения: {ex.Message}");
                        await Task.Delay(1000);
                        retryCount++;
                    }
                }
            }
        }

        public static void TGStart()
        {
            using (var workbookTG = new XLWorkbook(PathTG))
            {
                var sheetTG = workbookTG.Worksheet(1);

                if (!sheetTG.Cell(1, 2).IsEmpty() && !sheetTG.Cell(2, 2).IsEmpty())
                {
                    Token = sheetTG.Cell(1, 2).Value.ToString();
                    Chat.Id = Convert.ToInt64(sheetTG.Cell(2, 2).Value);
                    Client = new TelegramBotClient(Token);
                    Notify = sheetTG.Cell(3, 2).Value.ToString();
                    Sorting = Convert.ToInt32(sheetTG.Cell(4, 2).Value);
                    Buttons = sheetTG.Cell(5, 2).Value.ToString();
                    Bashorg = sheetTG.Cell(6, 2).Value.ToString();
                }
                else
                {
                    Console.WriteLine("Не указан Token или Id");
                }
            }
            SendMessageAsync("🤖 GridBoviBot подключен.\nНажмите /start чтобы включить ⌨").Wait();
        }

        public static async Task WaitMessage()
        {
            _botClient = new TelegramBotClient(Token);
            _receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = new[] { UpdateType.Message },
                ThrowPendingUpdates = true,
            };
            using var cts = new CancellationTokenSource();
            _botClient.StartReceiving(UpdateHandler, ErrorHandler, _receiverOptions, cts.Token);
        }

        private static async Task UpdateHandler(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            try
            {
                if (update.Type == UpdateType.Message && update.Message?.Type == MessageType.Text)
                {
                    var message = update.Message;
                    var chat = message.Chat;

                    switch (message.Text)
                    {
                        case "/start":
                            await SendReplyKeyboardAsync(chat.Id, "⌨️ Клавиатура подключена");
                            break;
                        case "📊 Отчет":
                            if (!string.IsNullOrEmpty(Report))
                            {
                                await botClient.SendTextMessageAsync(chat.Id, Report);
                            }
                            break;
                        case "💭 Цитаты Bash":
                            await BashOrgNet();
                            break;
                        case "💬 BOVI Флудилка":
                            await botClient.SendTextMessageAsync(chat.Id, "Перейдите по ссылке\n https://t.me/c/2046625015/1");
                            break;
                        default:
                            await botClient.SendTextMessageAsync(chat.Id, "Используй только кнопки!");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка в обработчике обновлений: {ex.Message}");
            }
        }

        private static Task ErrorHandler(ITelegramBotClient botClient, Exception error, CancellationToken cancellationToken)
        {
            var errorMessage = error switch
            {
                ApiRequestException apiRequestException => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
                _ => error.ToString()
            };

            Console.WriteLine(errorMessage);
            return Task.CompletedTask;
        }

        public static async Task BashOrgNet()
        {
            string content = await GetSecondQuote();
            if (content != null)
            {
                await SendMessageAsync(content);
            }
            else
            {
                Console.WriteLine("Не удалось найти содержимое.");
            }
        }

        private static async Task<string> GetSecondQuote()
        {
            try
            {
                var response = await HttpClient.GetStringAsync("https://bash.ru.net/random");
                var htmlDocument = new HtmlAgilityPack.HtmlDocument();
                htmlDocument.LoadHtml(response);

                var quoteNodes = htmlDocument.DocumentNode.SelectNodes("//div[@class='card-body']");
                if (quoteNodes != null && quoteNodes.Count > 1)
                {
                    string content = quoteNodes[1].InnerHtml;
                    content = content.Replace("<br>", Environment.NewLine)
                                     .Replace("&quot;", "\"")
                                     .Replace("&lt;", "<")
                                     .Replace("&gt;", ">");

                    content += Environment.NewLine + "Цитата с https://bash.ru.net/";

                    return content.Trim();
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Ошибка при получении цитаты: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Общая ошибка: {ex.Message}");
            }

            return null;
        }

        private static async Task SendReplyKeyboardAsync(long chatId, string message)
        {
            var replyKeyboard = new ReplyKeyboardMarkup(new[]
            {
                new KeyboardButton[] { "📊 Отчет", "💭 Цитаты Bash" },
                new KeyboardButton[] { "💬 BOVI Флудилка" }
            })
            {
                ResizeKeyboard = true,
            };
            await Client.SendTextMessageAsync(chatId, message, replyMarkup: replyKeyboard);
        }
    }
}
