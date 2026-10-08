using KPLN_BIMTools_Ribbon.Forms.Commands;
using KPLN_BIMTools_Ribbon.Forms.Models;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms.ViewModels
{
    public sealed class DBProjectCreateViewModel : INotifyPropertyChanged
    {
        private readonly DBProjectCreateWindow _window;
        private DBProjectWrapper _dbProjectWrapper;

        public DBProjectCreateViewModel(DBProjectCreateWindow window)
        {
            _window = window;
            DBPrjWrapper = new DBProjectWrapper();
            SetServerPathCommand = new RelayCommand(SetServerPath);
            CreateDBProjectCommand = new RelayCommand(CreateDBProject);
            CancelCommand = new RelayCommand(Cancel);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public ICommand SetServerPathCommand { get; }

        public ICommand CreateDBProjectCommand { get; }

        public ICommand CancelCommand { get; }

        public DBProjectWrapper DBPrjWrapper
        {
            get => _dbProjectWrapper;
            set
            {
                _dbProjectWrapper = value;
                OnPropertyChanged();
            }
        }

        private void SetServerPath()
        {
            using (System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (!string.IsNullOrEmpty(DBPrjWrapper.WrServerPath))
                    dialog.SelectedPath = DBPrjWrapper.WrServerPath;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK
                    && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    DBPrjWrapper.WrServerPath = dialog.SelectedPath;
                }
            }
        }

        private void CreateDBProject()
        {
            try
            {
                string normalizedMainPath = NormalizeServerPath(DBPrjWrapper.WrServerPath);
                if (!ValidateRequiredFields())
                    return;

                IEnumerable<DBProject> projects = SQLiteMainService
                    .SQLitePrjServiceInst
                    .GetDBProjects_ByRVersion(DBPrjWrapper.WrRevitVersion)
                    .ToArray();

                if (projects.Any(project =>
                    string.Equals(project.Code, DBPrjWrapper.WrCode, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(project.Stage, DBPrjWrapper.WrStage, StringComparison.OrdinalIgnoreCase)))
                {
                    ShowError($"Проект с кодом {DBPrjWrapper.WrCode} и стадией {DBPrjWrapper.WrStage} уже существует.");
                    return;
                }

                if (projects.Any(project => ProjectContainsAnyPath(project, normalizedMainPath)))
                {
                    ShowError("Проект по одному из указанных путей уже существует.");
                    return;
                }

                if (!ValidateMainPath() || !ValidateRevitServerPaths())
                    return;

                DBProject projectToCreate = new DBProject
                {
                    Name = DBPrjWrapper.WrName.Trim(),
                    Code = DBPrjWrapper.WrCode.Trim(),
                    Stage = DBPrjWrapper.WrStage,
                    RevitVersion = DBPrjWrapper.WrRevitVersion,
                    MainPath = normalizedMainPath,
                    RevitServerPath = DBPrjWrapper.WrRevitServerPath,
                    RevitServerPath2 = DBPrjWrapper.WrRevitServerPath2,
                    RevitServerPath3 = DBPrjWrapper.WrRevitServerPath3,
                    RevitServerPath4 = DBPrjWrapper.WrRevitServerPath4,
                    IsClosed = false
                };

                int createdProjectId = SQLiteMainService
                    .SQLitePrjServiceInst
                    .CreateDBDocument(projectToCreate)
                    .GetAwaiter()
                    .GetResult();

                if (createdProjectId <= 0)
                {
                    ShowError("База данных не вернула идентификатор созданного проекта.");
                    return;
                }

                MessageBox.Show(
                    _window,
                    $"Проект успешно создан. ID: {createdProjectId}",
                    "KPLN_DB: создание проекта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                _window.DialogResult = true;
                _window.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    _window,
                    $"При создании проекта возникла ошибка.\n\n{ex.Message}",
                    "KPLN_DB: ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private bool ValidateRequiredFields()
        {
            bool invalidCode = string.IsNullOrWhiteSpace(DBPrjWrapper.WrCode)
                || DBPrjWrapper.WrCode.Any(char.IsLower)
                || DBPrjWrapper.WrCode.Any(character => char.IsSeparator(character) || character == '~' || character == '/');

            bool invalidVersion = DBPrjWrapper.WrRevitVersion != 2020
                && DBPrjWrapper.WrRevitVersion != 2023
                && DBPrjWrapper.WrRevitVersion != 2024
                && DBPrjWrapper.WrRevitVersion != 2026;

            if (string.IsNullOrWhiteSpace(DBPrjWrapper.WrName)
                || invalidCode
                || string.IsNullOrWhiteSpace(DBPrjWrapper.WrStage)
                || invalidVersion
                || string.IsNullOrWhiteSpace(DBPrjWrapper.WrServerPath))
            {
                ShowError(
                    "Заполните имя, код, стадию, версию Revit и путь к папке стадии. " +
                    "Код должен состоять из заглавных букв, цифр и символов «.» или «_». " +
                    "Поддерживаются Revit 2020, 2023, 2024 и 2026.");
                return false;
            }

            return true;
        }

        private bool ValidateMainPath()
        {
            if (!Directory.Exists(DBPrjWrapper.WrServerPath))
            {
                ShowError("Указанной папки стадии не существует.");
                return false;
            }

            string directoryName = new System.IO.DirectoryInfo(DBPrjWrapper.WrServerPath).Name.ToLowerInvariant();
            if (!directoryName.Contains("стадия")
                && !directoryName.Contains("концепция")
                && !directoryName.Contains("агр")
                && !directoryName.Contains("аго"))
            {
                ShowError("Нужно указать корневую папку стадии: имя должно содержать «Стадия», «Концепция», «АГР» или «АГО».");
                return false;
            }

            return true;
        }

        private bool ValidateRevitServerPaths()
        {
            string[] paths =
            {
                DBPrjWrapper.WrRevitServerPath,
                DBPrjWrapper.WrRevitServerPath2,
                DBPrjWrapper.WrRevitServerPath3,
                DBPrjWrapper.WrRevitServerPath4
            };

            return paths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .All(path => !HasRevitServerPathError(path));
        }

        private bool HasRevitServerPathError(string path)
        {
            string[] parts = path.Split('/');
            if (parts.Length < 4 || string.IsNullOrWhiteSpace(parts[2]))
            {
                ShowError($"Путь Revit Server заполнен неверно: {path}");
                return true;
            }

            string hostName = parts[2];
            string rootDirectory = parts.Last(part => !string.IsNullOrWhiteSpace(part));

            try
            {
                RevitServer server = new RevitServer(hostName, DBPrjWrapper.WrRevitVersion);
                RevitServerAPILib.DirectoryInfo directoryInfo = server.GetDirectoryInfo(rootDirectory);
                if (directoryInfo.Exists)
                    return false;
            }
            catch (WebException ex)
            {
                ShowError($"Не удалось проверить путь Revit Server «{path}».\n{ex.Message}");
                return true;
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка работы с Revit Server «{path}».\n{ex.Message}");
                return true;
            }

            ShowError($"Путь Revit Server не существует: {path}");
            return true;
        }

        private bool ProjectContainsAnyPath(DBProject project, string normalizedMainPath)
        {
            return PathsEqual(project.MainPath, normalizedMainPath)
                || PathsEqual(project.RevitServerPath, DBPrjWrapper.WrRevitServerPath)
                || PathsEqual(project.RevitServerPath2, DBPrjWrapper.WrRevitServerPath2)
                || PathsEqual(project.RevitServerPath3, DBPrjWrapper.WrRevitServerPath3)
                || PathsEqual(project.RevitServerPath4, DBPrjWrapper.WrRevitServerPath4);
        }

        private static bool PathsEqual(string left, string right)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
                return false;

            return string.Equals(
                left.TrimEnd('\\', '/'),
                right.TrimEnd('\\', '/'),
                StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeServerPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return path;

            if (path.StartsWith(@"Y:\", StringComparison.OrdinalIgnoreCase))
                return @"\\stinproject.local\project\" + path.Substring(3);

            if (path.StartsWith(@"Z:\", StringComparison.OrdinalIgnoreCase))
                return @"\\fs01\lib\" + path.Substring(3);

            return path;
        }

        private void Cancel()
        {
            _window.DialogResult = false;
            _window.Close();
        }

        private void ShowError(string text)
        {
            MessageBox.Show(
                _window,
                text,
                "KPLN_DB: ошибка",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
