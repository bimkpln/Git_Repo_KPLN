using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPLN_Library_Bitrix24Worker
{
    /// <summary>
    /// Сервис по отправке сообщений в Битрикс
    /// </summary>
    public class BitrixMessageSender
    {
        private static UserDbService _userDbService;

        public static UserDbService CurrentUserDbService
        {
            get
            {
                if (_userDbService == null)
                    _userDbService = (UserDbService)new CreatorUserDbService().CreateService();

                return _userDbService;
            }
        }

        /// <summary>
        /// Отправить сообщение в чат бим-отдела
        /// </summary>
        /// <param name="msg">Сообщение, которое будет отправлено</param>
        public static async void SendMsg_ToBIMChat(string msg)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client
                        .GetAsync($"https://kpln.bitrix24.ru/rest/1310/b3zon6t9a38coxbq/im.message.add.json?MESSAGE={msg}&DIALOG_ID=chat99642");
                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(content))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Отправить сообщение в чат пользователя от имени спец. пользователя
        /// </summary>
        /// <param name="dBUser">Пользователь из БД КПЛН для отправки</param>
        /// <param name="msg">Сообщение, которое будет отправлено</param>
        public static async void SendMsg_ToUser_ByDBUser(DBUser dBUser, string msg)
        {
            int bitrixUserId = await GetDBUserBitrixId_ByDBUser(dBUser);
            if (bitrixUserId != -1)
            {
                try
                {
                    using (HttpClient client = new HttpClient())
                    {
                        // Выполнение GET - запроса к странице
                        HttpResponseMessage response = await client
                            .GetAsync($"https://kpln.bitrix24.ru/rest/1310/mfgd02ud186zswmp/im.message.add.json?MESSAGE={msg}&DIALOG_ID={bitrixUserId}");
                        if (response.IsSuccessStatusCode)
                        {
                            string content = await response.Content.ReadAsStringAsync();
                            if (string.IsNullOrEmpty(content))
                                throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Отправить сообщение по вебхуку и запросу в JSON
        /// </summary>
        /// <param name="dBUser">Пользователь из БД КПЛН для отправки</param>
        /// <param name="wh">Вебхук подготовленный в Битрикс</param>
        /// <param name="strJSON">Сообщение, которое будет отправлено в формате JSON</param>
        public static async void SendMsg_ToUser_ByWebhookKeyANDJSONRequest(string url, string json)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Наладжваем загалоўкі, каб паказаць, што запыт JSON
                    client.DefaultRequestHeaders.Add("Accept", "application/json");

                    // Падрыхтоўка кантэнту з JSON
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    // Адпраўка POST-запыту
                    HttpResponseMessage response = await client.PostAsync(url, content);
                    if (response.IsSuccessStatusCode)
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(responseContent))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Получить значение ID из Битрикс. Реализован формат заполнения значения ID для старых пользовтелей.
        /// </summary>
        /// <param name="dBUser">Пользователь из БД КПЛН для отправки</param>
        public static async Task<int> GetDBUserBitrixId_ByDBUser(DBUser dBUser)
        {
            if (dBUser.BitrixUserID != -1)
                return dBUser.BitrixUserID;

            return await SetDBUserBitrixId_ByDBUserSurname(dBUser);
        }

        /// <summary>
        /// Отправить данные по Id пользователя из Битрикс24 в БД КПЛН.
        /// </summary>
        /// <param name="dBUser">Пользователь из БД КПЛН для отправки</param>
        private static async Task<int> SetDBUserBitrixId_ByDBUserSurname(DBUser dBUser)
        {
            int id = -1;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client.GetAsync($"https://kpln.bitrix24.ru/rest/152/7nqwflagfu7wnirl/user.search.json?NAME={dBUser.Name}&LAST_NAME={dBUser.Surname}");
                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(content))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");

                        dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                        dynamic responseResult = dynDeserilazeData.result;
                        if (!int.TryParse(responseResult[0].ID.ToString(), out id))
                            throw new Exception("\n[KPLN]: Не удалось привести значение ID к int\n\n");
                    }
                }

                if (id == -1)
                    throw new Exception("\n[KPLN]: Ошибка получения пользователя из БД - не удалось получить id-пользователя Bitrix\n\n");

                CurrentUserDbService.UpdateDBUser_BitrixUserID(dBUser, id);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return id;
        }
    }
}
