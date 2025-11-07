using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace KPLN_ConsoleApp_Bitrix24Worker
{
    /// <summary>
    /// Программа удаления очистки от пользователей и удаление чатов. Одна проблема - нужны права сверх админа, или чтобы меня ставили владельцем чата
    /// </summary>
    internal class Program
    {
        private readonly static string _mainConfigPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_Config.json";
        private readonly static string _bitrixConfigs_MainWebHookName = "MainWebHook";
        private static string _webhookUrl;

        static async Task Main(string[] args)
        {
            // Подготовка
            Bitrix_Config[] bitrixConfigs = GetBitrixCongigs();

            _webhookUrl = bitrixConfigs.FirstOrDefault(d => d.Name == _bitrixConfigs_MainWebHookName).URL;

            // Основной блок
            Console.WriteLine("Введи имя чата для его очистки, или ID (если чатов много): ");
            string userInput = Console.ReadLine();
            Console.WriteLine();

            string chatId;
            if (int.TryParse(userInput, out int _))
                Console.WriteLine($"Оошибка ввода: Похоже, что это не id: \"{userInput}\". Нужно вводить ТОЛЬКО числа");
            else
            {
                chatId = await GetChatId_ByName(userInput);
                if (chatId != null)
                {
                    Console.WriteLine($"ID чата \"{userInput}\": {chatId}");

                    string[] usersIdArr = await GetUsersId_ByChatId(chatId);
                    if (usersIdArr != null)
                    {
                        Console.WriteLine($"В чате \"{userInput}\" - {usersIdArr.Length} пользователей/-ля");

                        Console.WriteLine();
                        Console.WriteLine("ВНИМАНИЕ: Эта процедура не обратима! Проверь корректность данных!!!");
                        Console.WriteLine("Удалить всех пользователей и сам чат? Да - Y, Нет - N : ");

                        string userDelAlg = Console.ReadLine();
                        if (userDelAlg.ToLower() == "y")
                        {
                            bool mainResult = await DeleteUser_ByChatIdAndUserIdArr(chatId, usersIdArr);
                            if (mainResult)
                                Console.WriteLine("Чат удален успешно!");
                            else
                                Console.WriteLine($"Скинь разработчику: Не удалось почистить чат {chatId}");
                        }
                        else
                            Console.WriteLine("Операция успешно отменена!");
                    }
                    else
                        Console.WriteLine($"Скинь разработчику: Ошибка получения пользователей из чата {chatId}");
                }
                else
                    Console.WriteLine($"Ошибка получения чата {userInput}. Сокрее всего - ошибка ввода имени чата (лучше скопировать его из браузера), либо вы не состоите в чате");

            }

            Console.ReadLine();
        }


        /// <summary>
        /// Поиск чата по имени (ВАЖНО: чаты могут быть с одинаковыми именами!)
        /// </summary>
        /// <param name="chatName">Имя чата</param>
        /// <returns>Если чатов несколько с таким именем - выдаёт пустое значение</returns>
        private static async Task<string> GetChatId_ByName(string chatName)
        {
            using (HttpClient client = new HttpClient())
            {
                HttpResponseMessage response = await client
                    .GetAsync($"{_webhookUrl}/im.search.chat.list.json?FIND={chatName}");
                string jsonResponse = await response.Content.ReadAsStringAsync();

                using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                {
                    JsonElement root = doc.RootElement;
                    JsonElement chats = root.GetProperty("result");

                    IEnumerable<JsonElement> chatNames = chats
                        .EnumerateArray()
                        .Where(chat => chat.GetProperty("name").GetString() == chatName);

                    if (chatNames.Count() > 1)
                    {
                        Console.WriteLine($"Чатов с именем {chatName} несколько. Нужно искать по ID");
                        return null;
                    }
                    else if (chatNames.Any())
                        return chatNames.FirstOrDefault().GetProperty("id").ToString();
                }
            }

            return null;
        }

        /// <summary>
        /// Получить пользователей из чата
        /// </summary>
        /// <param name="chatId"></param>
        /// <returns></returns>
        private static async Task<string[]> GetUsersId_ByChatId(string chatId)
        {
            using (HttpClient client = new HttpClient())
            {
                HttpResponseMessage response = await client
                    .GetAsync($"{_webhookUrl}/im.chat.user.list.json?CHAT_ID={chatId}");
                string jsonResponse = await response.Content.ReadAsStringAsync();

                using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                {
                    JsonElement root = doc.RootElement;
                    JsonElement users = root.GetProperty("result");
                    IEnumerable<JsonElement> usersId = users.EnumerateArray();
                    if (usersId.Any())
                        return usersId.Select(je => je.ToString()).ToArray();
                }
            }

            return null;
        }

        /// <summary>
        /// Удалить пользователей из чата
        /// </summary>
        /// <param name="chatId"></param>
        /// <returns></returns>
        private static async Task<bool> DeleteUser_ByChatIdAndUserIdArr(string chatId, string[] usersIdArr)
        {
            bool result = true;
            using (HttpClient client = new HttpClient())
            {
                // Удаляю всех
                foreach (string userId in usersIdArr)
                {
                    if (userId != "152")
                    {
                        HttpResponseMessage response = await client
                            .GetAsync($"{_webhookUrl}im.chat.user.delete.json?CHAT_ID={chatId}&USER_ID={userId}");
                        string jsonResponse = await response.Content.ReadAsStringAsync();
                        if (response.StatusCode == System.Net.HttpStatusCode.OK)
                        {
                            using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                            {
                                JsonElement root = doc.RootElement;
                                JsonElement respRes = root.GetProperty("result");

                                if (!respRes.GetBoolean())
                                    return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Очистить чат может только администратор");
                            return false;
                        }
                    }
                }

                // Удаляю админа
                string chatAdmin = usersIdArr.Where(id => id == "152").FirstOrDefault();
                if (chatAdmin != null)
                {
                    HttpResponseMessage response = await client
                            .GetAsync($"{_webhookUrl}im.chat.leave.json?CHAT_ID={chatId}");
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                        {
                            JsonElement root = doc.RootElement;
                            JsonElement respRes = root.GetProperty("result");

                            if (!respRes.GetBoolean())
                                return false;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Очистить чат может только администратор");
                        return false;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Получить коллекцию Bitrix_Config
        /// </summary>
        /// <returns></returns>
        private static Bitrix_Config[] GetBitrixCongigs()
        {
            string jsonConfig = File.ReadAllText(_mainConfigPath);
            JObject root = JObject.Parse(jsonConfig);

            var bitrixSection = root["BitrixConfig"]?["WEBHooks"];
            var bitrixList = new List<Bitrix_Config>();

            if (bitrixSection != null)
            {
                var bitrixObj = bitrixSection.ToObject<Bitrix_Config>();
                if (bitrixObj != null)
                    bitrixList.Add(bitrixObj);
            }

            return bitrixList.ToArray();
        }
    }
}
