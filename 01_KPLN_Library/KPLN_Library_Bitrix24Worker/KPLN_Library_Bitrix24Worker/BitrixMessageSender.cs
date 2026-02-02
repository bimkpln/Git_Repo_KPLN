using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Loader.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace KPLN_Library_Bitrix24Worker
{
    /// <summary>
    /// Сервис по отправке сообщений в Битрикс
    /// </summary>
    public class BitrixMessageSender
    {
        private static UserDbService _userDbService;
        /// <summary>
        /// Лучше заменить на ссылку на KPLN_Loader. Пока захаркодил, т.к. KPLN_Loader нужно всем переустановить (21.10.2025)
        /// </summary>
        private protected static readonly string _mainConfigPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_Config.json";
        /// <summary>
        /// Лучше заменить на ссылку на KPLN_Loader. Пока захаркодил, т.к. KPLN_Loader нужно всем переустановить (21.10.2025)
        /// </summary>
        private protected static readonly string _bitrixConfigs_MainWebHookName = "MainWebHook";
        private static Bitrix_Config[] _bitrixConfigs;

        /// <summary>
        /// Ссылка на пользователя
        /// </summary>
        private static UserDbService CurrentUserDbService
        {
            get
            {
                if (_userDbService == null)
                    _userDbService = (UserDbService)new CreatorUserDbService().CreateService();

                return _userDbService;
            }
        }

        /// <summary>
        /// Коллекция десерилизованныйх данных по настройкам Bitrix
        /// </summary>
        public static Bitrix_Config[] BitrixConfigs
        {
            get
            {
                if (_bitrixConfigs == null)
                    _bitrixConfigs = GetBitrixCongigs();

                return _bitrixConfigs;
            }
        }

        /// <summary>
        /// Ссылка на вебхук битрикс
        /// </summary>
        private static string WebHookUrl => BitrixConfigs.FirstOrDefault(conf => conf.Name == _bitrixConfigs_MainWebHookName).URL;

        #region Отправка сообщений
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
                    var requestData = new Dictionary<string, object>
                    {
                        { "MESSAGE",  msg},
                        { "DIALOG_ID", "chat99642"},
                    };

                    var jsonContent = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync($"{WebHookUrl}/im.message.add", jsonContent);
                    string responseContent = await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode)
                        throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Отправить сообщение в чат пользователей. Если чата нет - создать
        /// </summary>
        public static async Task<bool> SendMsg_ToUsersChat(DBUser dBUserSender, IEnumerable<DBUser> dBUsersReceiver, string msg, string imgId = "")
        {
            // Собираю данные по пользователям битрикс
            int bitrixUserIdSender = await GetDBUserBitrixId_ByDBUser(dBUserSender);
            List<int> bitrixUserIdsReceiver = new List<int>();
            foreach(var user in dBUsersReceiver)
            {
                bitrixUserIdsReceiver.Add(await GetDBUserBitrixId_ByDBUser(user));
            }
            
            if (bitrixUserIdSender == -1 || bitrixUserIdsReceiver.Any(id => id == -1)) 
                throw new Exception("\n[KPLN]: Ошибка получения ID пользователя Bitrix\n\n");


            // Формирую имя чата
            string chatTitle = string.Empty;
            string chatTitleRev = string.Empty;
            if (bitrixUserIdsReceiver.Count == 1)
            {
                chatTitle = $"{dBUserSender.Surname} {dBUserSender.Name} и {dBUsersReceiver.FirstOrDefault().Surname} {dBUsersReceiver.FirstOrDefault().Name}";
                chatTitleRev = $"{dBUsersReceiver.FirstOrDefault().Surname} {dBUsersReceiver.FirstOrDefault().Name} и {dBUserSender.Surname} {dBUserSender.Name}";
            }
            else
            {
                chatTitle = $"{dBUserSender.Surname} {dBUserSender.Name} и остальные";
                chatTitleRev = chatTitle;
            }


            // Работа с чатом
            List<int> finishIDListOfBitrUsers = bitrixUserIdsReceiver;
            // Добавляю робота бим, чтобы чат был
            finishIDListOfBitrUsers.Insert(0, 1310);
            finishIDListOfBitrUsers.Insert(1, bitrixUserIdSender);
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Ищу чат
                    string chatID = string.Empty;
                    HttpResponseMessage responseGetLastChats = await client.GetAsync($"{WebHookUrl}/im.recent.list");
                    if (!responseGetLastChats.IsSuccessStatusCode)
                        throw new Exception("\n[KPLN]: Ошибка получения чатов из Bitrix\n\n");

                    string responseContent = await responseGetLastChats.Content.ReadAsStringAsync();

                    dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);

                    foreach (var item in jsonResponse?.result.items)
                    {
                        var itemTitle = item?.title;
                        if (itemTitle == chatTitle || itemTitle == chatTitleRev)
                            chatID = item.id;
                    }

                        
                    // Чат не найден. Создаю
                    if (string.IsNullOrEmpty(chatID))
                    {
                        var requestNewChatData = new Dictionary<string, object>
                        {
                            ["TITLE"] = chatTitle,
                            ["USERS"] = finishIDListOfBitrUsers,
                        };

                        var jsonNeChatContent = new StringContent(JsonConvert.SerializeObject(requestNewChatData), Encoding.UTF8, "application/json");
                        var responseNewChat = await client.PostAsync($"{WebHookUrl}/im.chat.add", jsonNeChatContent);
                        string responseNewChatContent = await responseNewChat.Content.ReadAsStringAsync();
                        if (!responseNewChat.IsSuccessStatusCode)
                            throw new Exception("\n[KPLN]: Не удалось создать чат в Bitrix\n\n");

                        dynamic jsonNewChatResponse = JsonConvert.DeserializeObject(responseNewChatContent);

                        chatID = $"chat{jsonNewChatResponse?.result}";
                    }


                    // Отправляю в чат
                    StringContent contentMsg;
                    if (!string.IsNullOrEmpty(imgId))
                    {
                        HttpResponseMessage responseGetImg = await client.GetAsync($"{WebHookUrl}/disk.file.get?id={imgId}");
                        string responseGetImgContent = await responseGetImg.Content.ReadAsStringAsync();
                        dynamic jResp = JsonConvert.DeserializeObject(responseGetImgContent);

                        string detailUrl = jResp?.result?.DETAIL_URL; 
                        string downloadUrl = jResp?.result?.DOWNLOAD_URL;
                        string name = jResp?.result?.NAME;

                        var msgObj = new
                        {
                            DIALOG_ID = chatID,
                            MESSAGE = msg,
                            ATTACH = new[]
                            {
                                new
                                {
                                    IMAGE = new
                                    {
                                        NAME = name,
                                        LINK = detailUrl,
                                        PREVIEW = downloadUrl,
                                        WIDTH = 300,
                                        HEIGHT = 300
                                    }
                                }
                            }
                        };

                        contentMsg = new StringContent(
                            JsonConvert.SerializeObject(msgObj),
                            Encoding.UTF8,
                            "application/json"
                        );
                    }
                    else
                    {
                        var msgObj = new Dictionary<string, object>
                        {
                            ["DIALOG_ID"] = chatID,
                            ["MESSAGE"] = msg
                        };

                        contentMsg = new StringContent(
                            JsonConvert.SerializeObject(msgObj),
                            Encoding.UTF8,
                            "application/json"
                        );
                    }

                    // Отправка POST
                    HttpResponseMessage response = await client.PostAsync($"{WebHookUrl}/im.message.add", contentMsg);
                    if (!response.IsSuccessStatusCode)
                        throw new Exception("\n[KPLN]: Ошибка отправик сообщения в чат Bitrix\n\n");
                    
                    // Читаю итог
                    string responseMsgContent = await response.Content.ReadAsStringAsync();
                    if (string.IsNullOrEmpty(responseMsgContent))
                        throw new Exception("\n[KPLN]: Ошибка отправки сообщения в чат Bitrix\n\n");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Отправить сообщение в чат пользователя от имени спец. пользователя
        /// </summary>
        /// <param name="dBUser">Пользователь из БД КПЛН для отправки</param>
        /// <param name="msg">Сообщение, которое будет отправлено</param>
        /// <param name="system">Отображать сообщения в виде системного сообщения или нет ("Y" или "N")</param>
        public static async void SendMsg_ToUser_ByDBUser(DBUser dBUser, string msg, string system = "N")
        {
            int bitrixUserId = await GetDBUserBitrixId_ByDBUser(dBUser);
            if (bitrixUserId == -1) return;


            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var requestData = new Dictionary<string, object>
                    {
                        { "MESSAGE",  msg},
                        { "DIALOG_ID", bitrixUserId},
                        { "SYSTEM", system},
                    };

                    var jsonContent = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync($"{WebHookUrl}/im.message.add", jsonContent);
                    string responseContent = await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode)
                        throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Отправить сообщение по вебхуку и запросу в JSON
        /// </summary>
        /// <param name="dBUser">Пользователь из БД КПЛН для отправки</param>
        /// <param name="wh">Вебхук подготовленный в Битрикс</param>
        /// <param name="strJSON">Сообщение, которое будет отправлено в формате JSON</param>
        public static async void SendMsg_ToUser_ByJSONRequest(string json)
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
                    HttpResponseMessage response = await client.PostAsync($"{WebHookUrl}/im.message.add", content);
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
        /// Отправить комментарий в задачу
        /// </summary>
        /// <param name="taskId">ID задачи</param>
        /// <param name="msg">Сообщение</param>
        public static async Task<bool> SendMsgToTask_ByTaskId(int taskId, string msg)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var requestData = new Dictionary<string, string>
                    {
                        { "TASKID", taskId.ToString() },
                        { "FIELDS[POST_MESSAGE]", msg },
                    };

                    var content = new FormUrlEncodedContent(requestData);
                    var response = await client.PostAsync($"{WebHookUrl}/task.commentitem.add", content);
                    return response.IsSuccessStatusCode;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return false;
        }
        #endregion

        #region Работа с задачами
        /// <summary>
        /// Проверка задачи, на то, что она открыта: https://apidocs.bitrix24.ru/api-reference/tasks/fields.html
        /// </summary>
        /// <param name="taskId"></param>
        /// <returns>2 - Ждет выполнения, 3 - Выполняется, 4 - Ожидает контроля, 5 - Завершена, 6 - Отложена.</returns>
        public static async Task<bool> CheckTaskOpens_ByTaskId(int taskId)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/tasks.task.get.json?taskId={taskId}");

                    string content = await response.Content.ReadAsStringAsync();
                    if (string.IsNullOrEmpty(content))
                        throw new Exception("\n[KPLN]: Ошибка получения ответа при поиске задачи в Bitrix\n\n");
                    else
                    {
                        dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                        string strStatusId = dynDeserilazeData?.result?.task?.status?.Value;

                        if (int.TryParse(strStatusId, out int resultStaus))
                            return resultStaus == 2 || resultStaus == 3 || resultStaus == 6;
                        else
                            throw new Exception("\n[KPLN]: Ошибка получения статуса задачи Bitrix\n\n");
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return false;
        }


        /// <summary>
        /// Создать задачу по указанным полям. Наблюдатели заполняются автоматом
        /// </summary>
        /// <param name="groupId">ID группы в битрикс</param>
        /// <param name="title">Заголов</param>
        /// <param name="description">Описание задачи</param>
        /// <param name="groupId">ID группы Битрикс</param>
        /// <param name="parentTaskId">ID сборной задачи</param>
        /// <param name="tag">Тэг задачи для группирования</param>
        /// <param name="createdUserId">ID пользователя постановщика</param>
        /// <param name="respUserId">ID пользователя ответсвенного</param>
        public static async Task<string> CreateTask_ByMainFields_AutoAuditors(
            int groupId,
            string title,
            string description,
            int parentTaskId,
            string tag,
            int createdUserId,
            int respUserId)
        {
            // Получаю id руководителей постановщика и исполнителя
            int firstAuditorId = await GetUserHeadPersanBitrixId_ByUserId(createdUserId);
            int secondAuditorId = await GetUserHeadPersanBitrixId_ByUserId(respUserId);
            if (firstAuditorId == -1 || secondAuditorId == -1) return null;

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client
                        .GetAsync(
                            $"{WebHookUrl}/task.item.add.json?" +
                            $"FIELDS[TITLE]={title}" +
                            $"&FIELDS[DESCRIPTION]={description}" +
                            $"&FIELDS[GROUP_ID]={groupId}" +
                            $"&FIELDS[PARENT_ID]={parentTaskId}" +
                            $"&FIELDS[TAGS]={tag}" +
                            // Можно только если учётка админа. Пока тестирую формат постановки задач от робота
                            //$"&FIELDS[CREATED_BY]={createdUserId}" +
                            $"&FIELDS[RESPONSIBLE_ID]={respUserId}" +
                            $"&FIELDS[AUDITORS][]={firstAuditorId}" +
                            $"&FIELDS[AUDITORS][]={secondAuditorId}");
                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(content))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");
                        else
                        {
                            dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                            return dynDeserilazeData?.result?.ToString();
                        }
                    }
                    else
                        MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {response}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return null;
        }
        /// <summary>
        /// Обновить задачу - прикрепить файл
        /// </summary>
        /// <param name="taskId">Идентификатор задачи</param>
        /// <param name="imgBitrId">ИД файла с диска битрикс</param>
        public static async Task<bool> UpdateTask_LoadImg(int taskId, int imgBitrId)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var requestData = new Dictionary<string, string>
                    {
                        { "taskId", taskId.ToString() },
                        { "fileId", imgBitrId.ToString() }
                    };

                    var content = new FormUrlEncodedContent(requestData);
                    var response = await client.PostAsync($"{WebHookUrl}/tasks.task.files.attach", content);
                    return response.IsSuccessStatusCode;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return false;
        }

        /// <summary>
        /// Получить все подзадачи для родительской задчи
        /// </summary>
        /// <param name="parnetTaskId">Идентификатор родительской задачи</param>
        /// <returns>Словарь, где ключ - ID задачи из битриск; значение - ЗАГОЛОВО задачи из битрикс</returns>
        public static async Task<Dictionary<int, string>> GetAllSubTasks_IdAndTitle_ByParentId(int parnetTaskId)
        {
            Dictionary<int, string> result = new Dictionary<int, string>();
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var requestData = new Dictionary<string, object>
                    {
                        { "filter", new Dictionary<string, object> { { "PARENT_ID", parnetTaskId } } }
                    };

                    var jsonContent = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync($"{WebHookUrl}/tasks.task.list", jsonContent);
                    string responseContent = await response.Content.ReadAsStringAsync();

                    dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);
                    try
                    {
                        foreach (var item in jsonResponse?.result?.tasks)
                        {
                            var id = item.id?.Value;
                            if (!int.TryParse(id, out int taskId))
                                throw new Exception($"Ошибка парсинга {item.id} в int. Отправь разаботчику");
                            string taskTitle = item.title;

                            result[taskId] = taskTitle;
                        }
                    }
                    catch { }

                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return null;
        }

        #endregion

        #region Загрузка файлов на диск
        /// <summary>
        /// Загрузить файл на диск
        /// </summary>
        /// <returns>Id файла</returns>
        public static async Task<string> UploadFile_ToSpecialFolder(byte[] fileBytes, string fileName)
        {
            // Захожу в обший диск. ID диска = 11 (статичное поле, нет смысла его цеплять отдельно)
            string specailDictId = await GetDiskFolderID_ByRootIdANDName("11", "BIM_Файлы робота");

            if (string.IsNullOrEmpty(specailDictId))
            {
                MessageBox.Show(
                    "Ошибка: не удалось получить папку для загрузки файла.",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return null;
            }

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // 1. Атрымліваем uploadUrl
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/disk.folder.uploadfile?id={specailDictId}");
                    string responseContent = await response.Content.ReadAsStringAsync();

                    dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);
                    string uploadUrl = jsonResponse?.result?.uploadUrl;

                    if (string.IsNullOrEmpty(uploadUrl))
                    {
                        MessageBox.Show(
                            "Ошибка: не удалось получить URL для загрузки файла.",
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return null;
                    }

                    // 2. Загружаем файл на атрыманую спасылку
                    using (var fileContent = new ByteArrayContent(fileBytes))
                    {
                        fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png");

                        using (var multipartContent = new MultipartFormDataContent())
                        {
                            multipartContent.Add(fileContent, "file", $"{fileName}.png");

                            HttpResponseMessage response2 = await client.PostAsync(uploadUrl, multipartContent);
                            string responseContent2 = await response2.Content.ReadAsStringAsync();

                            dynamic jsonResponse2 = JsonConvert.DeserializeObject(responseContent2);
                            return jsonResponse2?.result?.ID;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке файла на диск Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Получить ID диска для группы
        /// </summary>
        /// <param name="groupId"></param>
        /// <returns></returns>
        public static async Task<string> GetDiskId_ByGroupId(string groupId)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/disk.storage.getlist");
                    if (response.IsSuccessStatusCode)
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();

                        dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);

                        foreach (var item in jsonResponse?.result)
                        {
                            var entityId = item?.ENTITY_ID;
                            if (entityId == groupId)
                                return item.ID;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return string.Empty;
        }

        /// <summary>
        /// Получить ID папки по имени в нужном корне
        /// </summary>
        /// <returns></returns>
        public static async Task<string> GetDiskFolderID_ByRootIdANDName(string rootId, string folderName)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/disk.storage.getchildren?id={rootId}");
                    if (response.IsSuccessStatusCode)
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();

                        dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);

                        foreach (var item in jsonResponse?.result)
                        {
                            var entityName = item?.NAME;
                            if (entityName == folderName)
                                return item.ID;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return string.Empty;
        }
        #endregion

        #region Дополнительные действия с пользователем Bitrix
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
        /// Получить Id проекта Bitrix по id задачи Bitrix
        /// </summary>
        /// <returns></returns>
        public static async Task<int> GetBitrixGroupId_ByTaskId(int taskId)
        {
            int groupId = -1;

            if (taskId == 0 || taskId == -1)
            {
                MessageBox.Show(
                    $"Ошибка при поиске группы в Bitrix - не удалось определить группу по ID указанной задачи",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/task.item.getdata.json?ID={taskId}");
                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(content))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");

                        dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                        dynamic responseResult = dynDeserilazeData.result;

                        var groupIdStr = responseResult.GROUP_ID.Value;
                        int.TryParse(groupIdStr, out groupId);
                    }
                }

                if (groupId == -1)
                    throw new Exception("\n[KPLN]: Ошибка получения группы Bitrix по указанной по id задаче Bitrix\n\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return groupId;
        }

        /// <summary>
        /// Получить Id пользователя-руководителя Bitrix
        /// </summary>
        /// <returns></returns>
        public static async Task<int> GetUserHeadPersanBitrixId_ByUserId(int currentUserId)
        {
            int headPersanId = -1;

            if (currentUserId == 0 || currentUserId == -1)
            {
                MessageBox.Show(
                    $"Ошибка при поиске руководителя для сотрудника - у пользователя НЕТ BitrixId в БД, либо возникла проблема с его поиском",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/user.search.json?ID={currentUserId}");
                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(content))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");

                        dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                        dynamic responseResult = dynDeserilazeData.result;
                        return await GetDepartmentHeadPersan_ByDepId((int)responseResult[0].UF_DEPARTMENT[0].Value);
                    }
                }

                if (headPersanId == -1)
                    throw new Exception("\n[KPLN]: Ошибка получения пользователя из БД - не удалось получить id-пользователя Bitrix\n\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return headPersanId;
        }

        /// <summary>
        /// Получить Id пользователя-руководителя отдела
        /// </summary>
        /// <returns></returns>
        public static async Task<int> GetDepartmentHeadPersan_ByDepId(int depId)
        {
            int headPersanId = -1;

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/department.get.json?ID={depId}");
                    if (response.IsSuccessStatusCode)
                    {
                        string content = await response.Content.ReadAsStringAsync();
                        if (string.IsNullOrEmpty(content))
                            throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");

                        dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                        dynamic responseResult = dynDeserilazeData.result;
                        var uf_Head = responseResult[0].UF_HEAD;
                        if (uf_Head == null)
                        {
                            var parent = responseResult[0].PARENT.Value;
                            if (int.TryParse(parent, out int parenDepId))
                                headPersanId = await GetDepartmentHeadPersan_ByDepId(parenDepId);
                        }
                        else
                            int.TryParse(responseResult[0].UF_HEAD.Value, out headPersanId);
                    }
                }

                if (headPersanId == -1)
                    throw new Exception("\n[KPLN]: Ошибка получения пользователя из БД - не удалось получить id-пользователя Bitrix\n\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке сообщения в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return headPersanId;
        }
        #endregion

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
                    HttpResponseMessage response = await client.GetAsync($"{WebHookUrl}/user.search.json?NAME={dBUser.Name}&LAST_NAME={dBUser.Surname}");
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

        /// <summary>
        /// Получить конфиги по Bitrix. Лучше заменить на ссылку на KPLN_Loader. Пока захаркодил, т.к. KPLN_Loader нужно всем переустановить (21.10.2025)
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
