using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
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
        private static readonly string _webHookUrl = "https://kpln.bitrix24.ru/rest/1310/uemokhg11u78vdvs";
        private static UserDbService _userDbService;

        private static UserDbService CurrentUserDbService
        {
            get
            {
                if (_userDbService == null)
                    _userDbService = (UserDbService)new CreatorUserDbService().CreateService();

                return _userDbService;
            }
        }

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
                    // Выполнение GET - запроса к странице
                    HttpResponseMessage response = await client
                        .GetAsync($"{_webHookUrl}/im.message.add.json?MESSAGE={msg}&DIALOG_ID=chat99642");
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
                            .GetAsync($"{_webHookUrl}/im.message.add.json?MESSAGE={msg}&DIALOG_ID={bitrixUserId}");
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
                    var response = await client.PostAsync($"{_webHookUrl}/task.commentitem.add", content);
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
                    HttpResponseMessage response = await client.GetAsync($"{_webHookUrl}/tasks.task.get.json?taskId={taskId}");

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
                            $"{_webHookUrl}/task.item.add.json?" +
                            $"FIELDS[TITLE]={title}" +
                            $"&FIELDS[DESCRIPTION]={description}" +
                            $"&FIELDS[GROUP_ID]={groupId}" +
                            $"&FIELDS[PARENT_ID]={parentTaskId}" +
                            $"&FIELDS[TAGS]={tag}" +
                            $"&FIELDS[CREATED_BY]={createdUserId}" +
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
                    var response = await client.PostAsync($"{_webHookUrl}/tasks.task.files.attach", content);
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
                    var response = await client.PostAsync($"{_webHookUrl}/tasks.task.list", jsonContent);
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
        /// <returns></returns>
        public static async Task<string> UploadFile(int groupId, byte[] fileBytes, string fileName)
        {
            int rootObjId = await GetDiskRootObjId_ByGroupId(groupId);

            try
            {
                using(HttpClient client = new HttpClient())
                {
                    // 1. Атрымліваем uploadUrl
                    HttpResponseMessage response = await client.GetAsync($"{_webHookUrl}/disk.folder.uploadfile?id={rootObjId}");
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
        public static async Task<int> GetDiskRootObjId_ByGroupId(int groupId)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync($"{_webHookUrl}/disk.storage.getlist");
                    if (response.IsSuccessStatusCode)
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();

                        dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);

                        foreach (var item in jsonResponse?.result)
                        {
                            var entityId = item?.ENTITY_ID;
                            if (entityId == $"{groupId}")
                                return item.ROOT_OBJECT_ID;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при постановке задачи в Bitrix: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Бывает, что не у всех проектов есть открытытй диск. В таком случае - кидаю на BIM (общая)
            return 23560;
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
                    HttpResponseMessage response = await client.GetAsync($"{_webHookUrl}/task.item.getdata.json?ID={taskId}");
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
                    HttpResponseMessage response = await client.GetAsync($"{_webHookUrl}/user.search.json?ID={currentUserId}");
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
                    HttpResponseMessage response = await client.GetAsync($"{_webHookUrl}/department.get.json?ID={depId}");
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
                    HttpResponseMessage response = await client.GetAsync($"{_webHookUrl}/user.search.json?NAME={dBUser.Name}&LAST_NAME={dBUser.Surname}");
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
