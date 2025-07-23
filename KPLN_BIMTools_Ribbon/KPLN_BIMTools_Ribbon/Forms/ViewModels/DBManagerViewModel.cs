using KPLN_BIMTools_Ribbon.Forms.Commands;
using KPLN_BIMTools_Ribbon.Forms.Models;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using RevitServerAPILib;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms.ViewModels
{
    public class DBManagerViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private DBProjectWrapper _createdDBProject;

        public ICommand SetServerPathCommand { get; }

        public ICommand CreateDBProjectCommand { get; }

        public DBManagerViewModel()
        {
            DBPrjWrapper = new DBProjectWrapper();

            SetServerPathCommand = new RelayCommand(SetServerPath);
            CreateDBProjectCommand = new RelayCommand(CreateDBProject);
        }

        public DBProjectWrapper DBPrjWrapper
        {
            get => _createdDBProject;
            set
            {
                _createdDBProject = value;
                OnPropertyChanged();
            }
        }

        private void SetServerPath()
        {
            using (System.Windows.Forms.FolderBrowserDialog openFolderDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (!string.IsNullOrEmpty(DBPrjWrapper.WrServerPath))
                    openFolderDialog.SelectedPath = DBPrjWrapper.WrServerPath;

                if (openFolderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK
                    && !string.IsNullOrWhiteSpace(openFolderDialog.SelectedPath))
                    DBPrjWrapper.WrServerPath = openFolderDialog.SelectedPath;
            }
        }

        private void CreateDBProject()
        {
            try
            {
                // Предварительная замена части пути к папке на серверное имя (так преобразуется ModelPath)
                string replacedServerPath = string.Empty;
                if (DBPrjWrapper.WrServerPath.StartsWith("Y:\\"))
                    replacedServerPath = DBPrjWrapper.WrServerPath.Replace("Y:\\", "\\\\stinproject.local\\project\\");
                else if (DBPrjWrapper.WrServerPath.StartsWith("Z:\\"))
                    replacedServerPath = DBPrjWrapper.WrServerPath.Replace("Y:\\", "\\\\fs01\\lib\\");

                #region Верификация
                // Проверка на пустые значения
                if (string.IsNullOrEmpty(DBPrjWrapper.WrName)
                    || (string.IsNullOrEmpty(DBPrjWrapper.WrCode) || DBPrjWrapper.WrCode.Any(c => char.IsLower(c)) || !DBPrjWrapper.WrCode.All(c => char.IsLetterOrDigit(c)))
                    || string.IsNullOrEmpty(DBPrjWrapper.WrStage)
                    || (DBPrjWrapper.WrRevitVersion != 2020 && DBPrjWrapper.WrRevitVersion != 2023)
                    || string.IsNullOrEmpty(DBPrjWrapper.WrServerPath))
                {
                    MessageBox.Show(
                        "Для создания проекта как минимум нужно указать:\n" +
                            "Имя проекта;\nКод проекта (только заглавные и цифры);\nСтадию проекта;\nВерсию Revit (2020 или 2023);\nПуть к папке стадии на сервере",
                        "KPLN_DB: Ошибка!",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }


                // Проверка на эквивалентные проекты в БД
                IEnumerable<DBProject> dbrjColl = DBMainService
                    .ProjectDbService
                    .GetDBProjects_ByRVersion(DBPrjWrapper.WrRevitVersion);

                IEnumerable<DBProject> dbrjColl_EqualCodeANDStage = dbrjColl
                    .Where(prj => prj.Code.Equals(DBPrjWrapper.WrCode) && prj.Stage.Equals(DBPrjWrapper.WrStage));
                if (dbrjColl_EqualCodeANDStage.Any())
                {
                    MessageBox.Show(
                        $"Проект с кодом {DBPrjWrapper.WrCode} и стадией {DBPrjWrapper.WrStage} - уже существует",
                        "KPLN_DB: Ошибка!",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }

                IEnumerable<DBProject> dbrjColl_EqualPath = dbrjColl
                    .Where(prj => prj.MainPath.Contains(replacedServerPath)
                        || (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath) && prj.RevitServerPath.Contains(DBPrjWrapper.WrRevitServerPath))
                        || (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath2) && prj.RevitServerPath.Contains(DBPrjWrapper.WrRevitServerPath2))
                        || (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath3) && prj.RevitServerPath.Contains(DBPrjWrapper.WrRevitServerPath3))
                        || (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath4) && prj.RevitServerPath.Contains(DBPrjWrapper.WrRevitServerPath4)));
                if (dbrjColl_EqualPath.Any())
                {
                    MessageBox.Show(
                        $"Проект по указанному пути - уже существует",
                        "KPLN_DB: Ошибка!",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }

                // Проверка существавания путей на серверах
                if (!Directory.Exists(DBPrjWrapper.WrServerPath))
                {
                    MessageBox.Show(
                        "Указанной папки не существует",
                        "KPLN_DB: Ошибка!",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }
                // Проверка что это папка стадии
                else
                {
                    string[] parts = DBPrjWrapper.WrServerPath.Split('\\');
                    if (parts.Length == 0)
                    {
                        MessageBox.Show(
                            "Невозможно проанализировать путь. Нужно использовать разделитеть '\'",
                            "KPLN_DB: Ошибка!",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);

                        return;
                    }

                    string lastPart = parts[parts.Length - 1];
                    if (!lastPart.ToLower().Contains("стадия") && !lastPart.ToLower().Contains("концепция") && !lastPart.ToLower().Contains("агр") && !lastPart.ToLower().Contains("аго"))
                    {
                        MessageBox.Show(
                            "Нужно указывать корневую папку стадии (имя должно содержать 'Концепция', или 'АГР', или 'АГО', или 'Cтадия')",
                            "KPLN_DB: Ошибка!",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);

                        return;
                    }
                }

                if (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath) && RSPathError(DBPrjWrapper.WrRevitServerPath))
                    return;

                if (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath2) && RSPathError(DBPrjWrapper.WrRevitServerPath2))
                    return;

                if (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath3) && RSPathError(DBPrjWrapper.WrRevitServerPath3))
                    return;

                if (!string.IsNullOrEmpty(DBPrjWrapper.WrRevitServerPath4) && RSPathError(DBPrjWrapper.WrRevitServerPath4))
                    return;
                #endregion

                // Создание
                Task<int> createTask = DBMainService.ProjectDbService.CreateDBDocument(new DBProject
                {
                    Name = DBPrjWrapper.WrName,
                    Code = DBPrjWrapper.WrCode,
                    Stage = DBPrjWrapper.WrStage,
                    RevitVersion = DBPrjWrapper.WrRevitVersion,
                    MainPath = replacedServerPath,
                    RevitServerPath = DBPrjWrapper.WrRevitServerPath,
                    RevitServerPath2 = DBPrjWrapper.WrRevitServerPath2,
                    RevitServerPath3 = DBPrjWrapper.WrRevitServerPath3,
                    RevitServerPath4 = DBPrjWrapper.WrRevitServerPath4,
                });
                createTask.Wait();

                if (createTask.Result > 0)
                {
                    MessageBox.Show(
                        "Проект успешно создан!",
                        "KPLN_DB: Создание проекта",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    $"При создании возникла ошибка. Для начала - убедись, что минимальный набор полей заполнен, и данные в них прошли верификацию (не горят красным)." +
                        $"\n\nЕсли всё с твоей стороны хорошо - отправь ошбику разработчику:\n{ex.Message}",
                    "KPLN_DB: Ошибка!",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private bool RSPathError(string rsPath)
        {
            string selectedRSHostName = rsPath.Split('/')[2];

            RevitServer revitServer = new RevitServer(selectedRSHostName, ModuleData.RevitVersion);
            string[] parts = rsPath.Split('/');
            string selectedRSMainDir = parts[parts.Length - 1];
            try
            {
                RevitServerAPILib.DirectoryInfo rsDirInfo = revitServer.GetDirectoryInfo(selectedRSMainDir);
                if (rsDirInfo.Exists)
                    return false;
            }
            catch (WebException wex)
            {
                if (wex.Message.Contains("Невозможно разрешить удаленное имя"))
                {
                    MessageBox.Show(
                        $"Указанного Revit-Server {rsPath} - не сущестует",
                        "KPLN_DB: Ошибка!",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
                else if (wex.Message.Contains("(404) Не найден"))
                {
                    MessageBox.Show(
                        $"Возможно, указанный путь на Revit-Server {rsPath} - не содержит корневой папки (это обязательное условие)",
                        "KPLN_DB: Ошибка!",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка работы с RS. Отправь разработчику:\n{ex.Message}",
                    "KPLN_DB: Ошибка!",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            return true;
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
