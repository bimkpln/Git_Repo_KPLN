using KPLN_CoordiantorAI.Common;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using WinForms = System.Windows.Forms;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class SettingsWindow : Window
    {
        private readonly CoordinatorAiRepository _repository;
        private Bitrix24Settings _bitrix24Settings;

        public SettingsWindow()
        {
            _repository = new CoordinatorAiRepository();
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            DbPathTextBox.Text = _repository.DatabaseFilePath;
            LoadGigaChatSettings();
            LoadBitrix24Settings();
            UpdateDatabaseStatus();
        }

        private void LoadGigaChatSettings()
        {
            GigaChatSettings settings = _repository.DatabaseExists
                ? _repository.LoadGigaChatSettings()
                : new GigaChatSettings();

            AuthUrlTextBox.Text = settings.AuthUrl;
            ApiUrlTextBox.Text = settings.ApiUrl;
            ClientIdTextBox.Text = settings.ClientId;
            ClientSecretPasswordBox.Password = settings.ClientSecret;
            ScopeTextBox.Text = settings.Scope;
            CertificatePathTextBox.Text = settings.CertificatePath;
            EmbeddingFolderPathTextBox.Text = settings.EmbeddingFolderPath;
            SystemPromptTextBox.Text = settings.SystemPrompt;
        }

        private void LoadBitrix24Settings()
        {
            _bitrix24Settings = _repository.DatabaseExists
                ? _repository.LoadBitrix24Settings()
                : CreateDefaultBitrix24Settings();

            BitrixWebhookUrlTextBox.Text = _bitrix24Settings.WebhookUrl;
            BitrixMessageModeComboBox.SelectedIndex = _bitrix24Settings.CoordinatorMessageMode == Bitrix24CoordinatorMessageMode.FirstQuestion ? 0 : 1;
            BitrixDepartmentsItemsControl.ItemsSource = _bitrix24Settings.DepartmentCoordinators;
        }

        private static Bitrix24Settings CreateDefaultBitrix24Settings()
        {
            Bitrix24Settings settings = new Bitrix24Settings();
            foreach (SubDepartmentInfo subDepartment in SubDepartmentNameResolver.GetKnownSubDepartments())
            {
                settings.DepartmentCoordinators.Add(new Bitrix24DepartmentCoordinator
                {
                    DepartmentId = subDepartment.Id,
                    DepartmentName = subDepartment.Name
                });
            }

            return settings;
        }

        private void OnCreateDatabaseClick(object sender, RoutedEventArgs e)
        {
            try
            {
                _repository.EnsureDatabase();
                LoadGigaChatSettings();
                LoadBitrix24Settings();
                UpdateDatabaseStatus();
                MessageBox.Show(this, "БД создана или обновлена.", "Координатор ИИ", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Ошибка создания БД", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnSaveClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!EnsureDatabaseAvailable())
                    return;

                _repository.SaveGigaChatSettings(new GigaChatSettings
                {
                    AuthUrl = AuthUrlTextBox.Text,
                    ApiUrl = ApiUrlTextBox.Text,
                    ClientId = ClientIdTextBox.Text,
                    ClientSecret = ClientSecretPasswordBox.Password,
                    Scope = ScopeTextBox.Text,
                    CertificatePath = CertificatePathTextBox.Text,
                    EmbeddingFolderPath = EmbeddingFolderPathTextBox.Text,
                    SystemPrompt = SystemPromptTextBox.Text
                });

                if (_bitrix24Settings == null)
                    _bitrix24Settings = CreateDefaultBitrix24Settings();

                _bitrix24Settings.WebhookUrl = BitrixWebhookUrlTextBox.Text;
                _bitrix24Settings.CoordinatorMessageMode = BitrixMessageModeComboBox.SelectedIndex == 0
                    ? Bitrix24CoordinatorMessageMode.FirstQuestion
                    : Bitrix24CoordinatorMessageMode.FullChat;
                _repository.SaveBitrix24Settings(_bitrix24Settings);

                UpdateDatabaseStatus();
                MessageBox.Show(this, "Настройки сохранены.", "Координатор ИИ", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Ошибка сохранения настроек", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OnBrowseCertificateClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Выберите сертификат GigaChat",
                CheckFileExists = true,
                Filter = "Сертификаты (*.cer;*.crt;*.pem)|*.cer;*.crt;*.pem|Все файлы (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) == true)
                CertificatePathTextBox.Text = dialog.FileName;
        }

        private void OnBrowseDatabaseFileClick(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                Title = "Выберите файл БД",
                AddExtension = true,
                DefaultExt = ".db",
                FileName = Path.GetFileName(_repository.DatabaseFilePath),
                Filter = "База данных SQLite (*.db)|*.db|Все файлы (*.*)|*.*",
                OverwritePrompt = false
            };

            if (Directory.Exists(_repository.DatabaseFolder))
                dialog.InitialDirectory = _repository.DatabaseFolder;

            if (dialog.ShowDialog(this) != true)
                return;

            try
            {
                CoordinatorAiDatabaseConfig.SetDatabaseFilePath(dialog.FileName);
                DbPathTextBox.Text = _repository.DatabaseFilePath;
                UpdateDatabaseStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Ошибка выбора файла БД", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnBrowseEmbeddingFolderClick(object sender, RoutedEventArgs e)
        {
            using (WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "Выберите папку с эмбеддингами";
                dialog.ShowNewFolderButton = true;

                string currentPath = (EmbeddingFolderPathTextBox.Text ?? string.Empty).Trim();
                if (Directory.Exists(currentPath))
                    dialog.SelectedPath = currentPath;

                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                    EmbeddingFolderPathTextBox.Text = dialog.SelectedPath;
            }
        }

        private void OnOpenDatabaseFolderClick(object sender, RoutedEventArgs e)
        {
            OpenDirectory(_repository.DatabaseFolder);
        }

        private void OnOpenEmbeddingFolderClick(object sender, RoutedEventArgs e)
        {
            OpenDirectory(EmbeddingFolderPathTextBox.Text);
        }

        private void OnOpenCertificateFolderClick(object sender, RoutedEventArgs e)
        {
            OpenContainingDirectory(CertificatePathTextBox.Text);
        }

        private void OnDeleteAllQuestionsClick(object sender, RoutedEventArgs e)
        {
            if (!_repository.DatabaseExists)
            {
                MessageBox.Show(this, "БД не найдена.", "Координатор ИИ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            MessageBoxResult result = MessageBox.Show(
                this,
                "Вы точно хотите удалить все вопросы пользователей? Это действие нельзя отменить.",
                "Координатор ИИ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                _repository.DeleteAllQuestionSessions();
                MessageBox.Show(this, "Вопросы пользователей удалены.", "Координатор ИИ", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Ошибка удаления вопросов", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnAddBitrixCoordinatorClick(object sender, RoutedEventArgs e)
        {
            Bitrix24DepartmentCoordinator department = (sender as FrameworkElement)?.Tag as Bitrix24DepartmentCoordinator;
            if (department == null)
                return;

            if (department.Coordinators.Count >= 3)
            {
                MessageBox.Show(this, "Для отдела можно добавить не больше трех ответственных.", "Bitrix24", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            BitrixCoordinatorDialog dialog = new BitrixCoordinatorDialog
            {
                Owner = this
            };

            if (dialog.ShowDialog() != true)
                return;

            department.Coordinators.Add(new Bitrix24CoordinatorContact
            {
                DepartmentId = department.DepartmentId,
                DepartmentName = department.DepartmentName,
                UserId = dialog.UserId,
                UserName = dialog.UserName
            });
        }

        private void OnRemoveBitrixCoordinatorClick(object sender, RoutedEventArgs e)
        {
            e.Handled = true;

            Bitrix24CoordinatorContact contact = (sender as FrameworkElement)?.Tag as Bitrix24CoordinatorContact;
            if (contact == null || _bitrix24Settings == null)
                return;

            foreach (Bitrix24DepartmentCoordinator department in _bitrix24Settings.DepartmentCoordinators)
            {
                if (department.Coordinators.Remove(contact))
                    return;
            }
        }

        private bool EnsureDatabaseAvailable()
        {
            if (_repository.DatabaseExists)
            {
                _repository.EnsureDatabase();
                return true;
            }

            MessageBoxResult result = MessageBox.Show(
                this,
                "БД не найдена. Создать новую?",
                "Координатор ИИ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return false;

            _repository.EnsureDatabase();
            return true;
        }

        private void UpdateDatabaseStatus()
        {
            DatabaseStatusTextBlock.Text = _repository.DatabaseExists
                ? "БД найдена: " + _repository.DatabaseFilePath
                : "БД не найдена: " + _repository.DatabaseFilePath;
        }

        private void OpenContainingDirectory(string filePath)
        {
            string directoryPath = Path.GetDirectoryName((filePath ?? string.Empty).Trim());
            OpenDirectory(directoryPath);
        }

        private void OpenDirectory(string directoryPath)
        {
            string normalizedPath = (directoryPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedPath) || !Directory.Exists(normalizedPath))
            {
                MessageBox.Show(this, "Папка не найдена.", "Координатор ИИ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                Process.Start("explorer.exe", normalizedPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Ошибка открытия папки", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}