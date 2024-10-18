using ClosedXML.Excel;
using ClosedXML.Report.Options;
using DocumentFormat.OpenXml.Spreadsheet;
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
        #region Переменные
        public static int SendReport { get; set; } = 50; // Если значение не указано, то отправлять отчет через 50 циклов Buy/Sell
        public static int Sorting { get; set; } = 50; // Если значение не указано, то сортировать монеты через 50 циклов Buy/Sell
        public static string Token { get; set; } = "";
        public static string Notify { get; set; } = "";
        public static string Bashorg { get; set; } = "";
        public static string Buttons { get; set; } = "";
        public static string Report = "📊 Подготовка отчета";
        public static TelegramBotClient Client = new(Token);
        public static Chat Chat = new();
        private static readonly string PathTG = @"..\\..\\..\\..\\Telegram.xlsx"; // Расположение конфигурационного файла
        private static ITelegramBotClient _botClient;
        private static ReceiverOptions _receiverOptions;
        private static readonly HttpClient HttpClient = new();
        #endregion

        #region Отправка сообщений
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
        #endregion

        #region Чтение параметров из Telegram.xlsx
        public static async Task TGConfig()
        {
            using (var workbookTG = new XLWorkbook(PathTG))
            {
                var sheetTG = workbookTG.Worksheet(1);

                if (!sheetTG.Cell(1, 2).IsEmpty() && !sheetTG.Cell(2, 2).IsEmpty())
                {
                    Token = sheetTG.Cell(1, 2).GetString();
                    Chat.Id = Convert.ToInt64(sheetTG.Cell(2, 2).Value);
                    Client = new TelegramBotClient(Token);
                    Notify = sheetTG.Cell(3, 2).GetString();
                    SendReport = Convert.ToInt32(sheetTG.Cell(4, 2).Value);
                    Buttons = sheetTG.Cell(5, 2).GetString();
                    Bashorg = sheetTG.Cell(6, 2).GetString();
                    Sorting = Convert.ToInt32(sheetTG.Cell(7, 2).Value);
                }
                else
                {
                    Console.WriteLine("Не указан Token или Id");
                }
            }
            SendMessageAsync("🤖 GridBoviBot подключен.\nНажмите /start чтобы включить ⌨").Wait();

            // Подключение телеграм кнопок
            if (TG.Buttons == "True")
            {
                await TG.WaitMessage();
            }
        }
        #endregion

        #region Ожидание сообщений от пользователя
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
        #endregion

        #region Получение и обработка цитат с BashOrgNet
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
        #endregion

        #region Подключение кнопок и обработка сообщений телеграм
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
        #endregion
    }
}
