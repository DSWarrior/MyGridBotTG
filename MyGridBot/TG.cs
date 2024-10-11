using ClosedXML.Excel;
using ClosedXML.Report.Options;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
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
        public static Int32 sorting = 50;
        public static string token = "";
        public static string notify = "";
        public static string buttons = "";
        public static TelegramBotClient client = new(token);
        public static Chat chat = new();
        //TGStart
        static string _pathTG = @"..\\..\\..\\..\\Telegram.xlsx";
        //WaitMessage
        public static ITelegramBotClient _botClient; // Это клиент для работы с Telegram Bot API, который позволяет отправлять сообщения, управлять ботом, подписываться на обновления и многое другое.
        public static ReceiverOptions _receiverOptions; // Это объект с настройками работы бота. Здесь мы будем указывать, какие типы Update мы будем получать, Timeout бота и так далее.
        public static string Report = ""; //"📊 Подготовка отчета";

        public static async Task Message(string message)
        {
            if (token != "")
            {
                int flag = 0;
                while (true)
                {
                    try
                    {
                        await client.SendTextMessageAsync(chat.Id, message);
                        break;
                    }
                    catch
                    {
                        Console.WriteLine("Не верно указан Token или IdChat.");
                        await Task.Delay(1000);
                        flag++;
                        if (flag == 1)
                        {
                            break;
                        }
                    }
                }
            }
        }
        public static void TGStart()
        {
            while (true)
            {
                using (var workbookTG = new XLWorkbook(_pathTG))
                {
                    var sheetTG = workbookTG.Worksheet(1);

                    if (!sheetTG.Cell(1, 2).IsEmpty() && !sheetTG.Cell(2, 2).IsEmpty())
                    {
                        token = sheetTG.Cell(1, 2).Value.ToString();
                        chat.Id = Convert.ToInt64(sheetTG.Cell(2, 2).Value);
                        client = new TelegramBotClient(token);
                        notify = sheetTG.Cell(3, 2).Value.ToString();
                        sorting = Convert.ToInt32(sheetTG.Cell(4, 2).Value);
                        buttons = sheetTG.Cell(5, 2).Value.ToString();
                    }
                    else
                    {
                        Console.WriteLine("Не указан Token или Id");
                    }
                    break;
                }
            }
            Message("🤖 Grid Bovi Bot подключен.");
        }
        public static async Task WaitMessage()
        {
            _botClient = new TelegramBotClient(token);
            _receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = new[] // Тут указываем типы получаемых Update`ов, о них подробнее расказано тут https://core.telegram.org/bots/api#update
            {
                UpdateType.Message,
            },
                ThrowPendingUpdates = true,
            };
            using var cts = new CancellationTokenSource();
            _botClient.StartReceiving(UpdateHandler, ErrorHandler, _receiverOptions, cts.Token); // Запускаем бота
        }
        private static async Task UpdateHandler(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            try
            {
                switch (update.Type)
                {
                    case UpdateType.Message:
                        {
                            var message = update.Message;
                            var user = message.From;
                            var chat = message.Chat;
                            switch (message.Type)
                            {
                                case MessageType.Text:
                                    {
                                        if (message.Text == "/start")
                                        {
                                            var replyKeyboard = new ReplyKeyboardMarkup(
                                                new List<KeyboardButton[]>()
                                                {
                                                    new KeyboardButton[]
                                                {
                                                    new KeyboardButton("📊 Отчет"),
                                                    new KeyboardButton("💬 BOVI Флудилка")
                                                }
                                                })
                                            {
                                                ResizeKeyboard = true,
                                            };
                                            await botClient.SendTextMessageAsync(chat.Id, "⌨️ Клавиатура подключена", replyMarkup: replyKeyboard);
                                            return;
                                        }
                                        if (message.Text == "📊 Отчет")
                                        {
                                            if (Report != "")
                                            {
                                                await botClient.SendTextMessageAsync(chat.Id, Report);
                                            }
                                            return;
                                        }
                                        if (message.Text == "💬 BOVI Флудилка")
                                        {
                                            await botClient.SendTextMessageAsync(chat.Id, "Перейдите по ссылке\n https://t.me/c/2046625015/1");
                                            return;
                                        }
                                        return;
                                    }
                                default:
                                    {
                                        await botClient.SendTextMessageAsync(
                                            chat.Id,
                                            "Используй только текст!");
                                        return;
                                    }
                            }
                        }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private static Task ErrorHandler(ITelegramBotClient botClient, Exception error, CancellationToken cancellationToken)
        {
            // Тут создадим переменную, в которую поместим код ошибки и её сообщение 
            var ErrorMessage = error switch
            {
                ApiRequestException apiRequestException
                    => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
                _ => error.ToString()
            };

            Console.WriteLine(ErrorMessage);
            return Task.CompletedTask;
        }
    }
}
