using Autodesk.Revit.UI;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using KPLN_Loader.Core.Entities;
using KPLN_Loader.Forms;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KPLN_Loader.Services
{
    /// <summary>
    /// Сервис для работы с google tabs
    /// </summary>
    internal sealed class GTabService
    {
        private readonly Logger _logger;
        private readonly SheetsService _gtabService;
        private readonly string _spreadsheetId;

        private readonly string _userSheetRange = "Users!A1:N";
        private readonly string _subDepsSheetRange = "SubDepartments!A1:D";

        ///<summary>
        /// Путь до локальной папки пользователя с конфигом и установленными файлами
        ///</summary>
        private readonly DirectoryInfo _mainLocation;

        private IEnumerable<SubDepartment> _subDepartments;

        internal GTabService(Logger logger, string spreadsheetId, string revitVersion)
        {
            _logger = logger;
            _spreadsheetId = spreadsheetId;

            _mainLocation = new DirectoryInfo(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                @"AppData\Roaming\Autodesk\Revit\Addins",
                $"{revitVersion}\\KPLN_Loader"));

            GoogleCredential credential;
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            using (var stream = new FileStream(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString() + "\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            _gtabService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Google Sheets Client"
            });
        }

        /// <summary>
        /// Кэширование коллекции отделов из БД
        /// </summary>
        private IEnumerable<SubDepartment> SubDepartments
        {
            get
            {
                if (_subDepartments == null)
                {
                    _subDepartments = GetEntityFromRange<SubDepartment>(_subDepsSheetRange);
                }

                return _subDepartments;
            }
        }

        #region Create
        /// <summary>
        /// Авторизация пользователя
        /// </summary>
        /// <returns>Текущий пользователь</returns>
        internal User Authorization()
        {
            Task<Guid> getUserGuid = Task.Run(() => GetUserGuidFromSystem());

            string currentDate = DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
            string sysUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();

            Task.WaitAll(getUserGuid);
            Guid systemUserGuid = getUserGuid.Result;

            User currentUser = GetEntityFromRange<User>(_userSheetRange).FirstOrDefault(user => user.SystemName.Equals(sysUserName) && user.SystemGuid.Equals(systemUserGuid.ToString()));
            if (currentUser == null)
            {
                LoginForm loginForm = new LoginForm(SubDepartments.Where(s => s.IsAuthEnabled), true);
                if ((bool)loginForm.ShowDialog())
                {
                    int userId = GetLastUserId() + 1;
                    currentUser = new User()
                    {
                        Id = userId,
                        SystemName = sysUserName,
                        SystemGuid = systemUserGuid.ToString(),
                        Name = loginForm.CreatedWPFUser.Name,
                        Surname = loginForm.CreatedWPFUser.Surname,
                        Company = loginForm.CreatedWPFUser.Company,
                        SubDepartmentId = loginForm.CreatedWPFUser.SubDepartment.Id,
                        RegistrationDate = currentDate,
                        IsUserRestricted = false,
                        IsDebugMode = false,
                        BitrixUserID = -10,
                        IsExtraNet = true,
                    };

                    AppendEntityToRange(_userSheetRange, currentUser);

                    _logger.Info($"Пользователь {currentUser.SystemName}: " +
                        $"{currentUser.Surname} {currentUser.Name} из отдела {GetSubDepartmentForCurrentUser(currentUser).Code} " +
                        $"успешно создан и записан в БД!");
                }
                else
                {
                    TaskDialog.Show(
                        "Ошибка",
                        "Пользователь отменил ввод данных. Загрузка завершена с ошибкой. Перезапустите Revit и заполните все строки в окне регистрации",
                        TaskDialogCommonButtons.Cancel);

                    return null;
                }
            }
            else
            {
                _logger.Info($"Пользователь {currentUser.SystemName}: " +
                    $"{currentUser.Surname} {currentUser.Name} из отдела {GetSubDepartmentForCurrentUser(currentUser).Code} " +
                    $"успешно определен!");
            }

            if (currentUser.IsDebugMode)
                _logger.Info("ВЫБРАН СТАТУС ЗАПУСКА - DEBUG");

            currentUser.LastConnectionDate = currentDate;
            
            string currentUserRange = FindEntityRangeByEntityKey(_userSheetRange, currentUser.SystemGuid, 3);
            WriteEntityToRange(currentUserRange, currentUser);

            return currentUser;
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить SubDepartment пользоватлея
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        internal SubDepartment GetSubDepartmentForCurrentUser(User currentUser) => SubDepartments.FirstOrDefault(s => s.Id == currentUser.SubDepartmentId);

        #endregion

        #region Update
        /// <summary>
        /// Обновление имени Revit-пользователя
        /// </summary>
        /// <param name="userName">Текущее имя в Ревит</param>
        /// <param name="currentUser">Пользователь для проверки из БД</param>
        internal void SetRevitUserName(string userName, User currentUser)
        {
            if (currentUser.RevitUserName == null || !currentUser.RevitUserName.Equals(userName))
            {
                // Меняю объект
                currentUser.RevitUserName = userName;
                // Записываю в таблицу
                string currentUserRange = FindEntityRangeByEntityKey(_userSheetRange, currentUser.SystemGuid, 3);
                WriteEntityToRange(currentUserRange, currentUser);
            }
        }
        #endregion

        /// <summary>
        /// Получить id последнего пользователя
        /// </summary>
        private int GetLastUserId()
        {
            User lastUser = GetEntityFromRange<User>(_userSheetRange).LastOrDefault();
            if (lastUser == null)
                return 0;

            return lastUser.Id;
        }

        /// <summary>
        /// Получить системный Guid пользователя (генриться автоматом, потом проверяется)
        /// </summary>
        private Guid GetUserGuidFromSystem()
        {
            Guid guid;

            string guidFileFullPath = Path.Combine(_mainLocation.FullName, "UserGuid.txt");
            if (!File.Exists(guidFileFullPath))
            {
                try
                {
                    DirectoryInfo dir = Directory.CreateDirectory(_mainLocation.FullName);
                    using (StreamWriter sw = File.CreateText(guidFileFullPath))
                    {
                        guid = Guid.NewGuid();
                        sw.Write(guid.ToString());
                    }

                    return guid;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Не удалось создать уникальный номер пользователя (Guid). Ошибка: {ex.Message}");
                }
            }

            using (StreamReader sr = File.OpenText(guidFileFullPath))
            {
                string guidLine = sr.ReadToEnd();
                Guid.TryParse(guidLine, out guid);
            }

            if (guid == null)
                throw new Exception("Не удалось получить уникальный номер пользователя (Guid)");

            return guid;
        }

        /// <summary>
        /// Получить определненную сущность из предоставленного диапозона
        /// </summary>
        private IEnumerable<T> GetEntityFromRange<T>(string range) where T : new()
        {
            var values = ReadRange(range);
            var properties = typeof(T).GetProperties();

            // Иду по строкам, за исключением 1 (шапка)
            foreach (var row in values.Skip(1))
            {
                var instance = new T();
                for (int i = 0; i < row.Count && i < properties.Length; i++)
                {
                    var property = properties[i];
                    if (property.CanWrite)
                    {
                        var value = row[i];
                        property.SetValue(instance, Convert.ChangeType(value, property.PropertyType));
                    }
                }

                yield return instance;
            }
        }

        private IList<IList<object>> ReadRange(string range)
        {
            var request = _gtabService.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = request.Execute();

            return response.Values ?? new List<IList<object>>();
        }

        /// <summary>
        /// Добавить сущность в конец строк указанного диапазноа
        /// </summary>
        private void AppendEntityToRange<T>(string range, T data)
        {
            var properties = typeof(T).GetProperties();
            var values = new List<IList<object>>();

            var row = properties.Select(p => p.GetValue(data)?.ToString()).ToList<object>();
            values.Add(row);

            AppendRange(range, values);
        }

        private void AppendRange(string range, IList<IList<object>> values)
        {
            var body = new ValueRange { Values = values };
            var request = _gtabService.Spreadsheets.Values.Append(body, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

            request.Execute();
        }

        /// <summary>
        /// Получить диапазон ячеек указанной сущности по ключу
        /// </summary>
        /// <param name="startRange">Стартовый диапазон для поиска</param>
        /// <param name="entityKey">Ключевое значение сущности для поиска</param>
        /// <param name="keyColumn">Номер столбца с ключом</param>
        private string FindEntityRangeByEntityKey(string startRange, string entityKey, int keyColumn)
        {
            // Предварительная проверка запроса
            string[] startSheetNameAndRange = startRange.Split(new char[] { '!', ':' });
            if (startSheetNameAndRange.Length != 3)
                throw new Exception("Нарушена структура запроса (ИмяЛиста!Столбец:Строка)");
            string columnName = Regex.Match(startSheetNameAndRange[1], @"[A-Z]+").Value;

            // Указываем диапазон для поиска
            var request = _gtabService.Spreadsheets.Values.Get(_spreadsheetId, startRange);
            var response = request.Execute();

            // Проверяем наличие значений
            var values = response.Values;
            if (values == null || values.Count == 0)
                throw new Exception("Таблица пустая или данные отсутствуют.");

            // Ищем сущность по ключу
            for (int i = 0; i < values.Count; i++)
            {
                var row = values[i];
                if (row[keyColumn - 1].ToString().Equals(entityKey, StringComparison.OrdinalIgnoreCase))
                {
                    // +1 для учета нумерации таблицы
                    int rowNumber = i + 1;

                    return $"{startSheetNameAndRange[0]}!{columnName}{rowNumber}:{startSheetNameAndRange[2]}";
                }
            }

            throw new Exception("Сущность не найдена.");
        }

        /// <summary>
        /// Перезаписать сущность в данном диапазоне
        /// </summary>
        private void WriteEntityToRange<T>(string range, T data)
        {
            var properties = typeof(T).GetProperties();
            var values = new List<IList<object>>();

            var row = properties.Select(p => p.GetValue(data)?.ToString()).ToList<object>();
            values.Add(row);

            WriteRange(range, values);
        }

        private void WriteRange(string range, IList<IList<object>> values)
        {
            var body = new ValueRange { Values = values };
            var request = _gtabService.Spreadsheets.Values.Update(body, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            request.Execute();
        }
    }
}
