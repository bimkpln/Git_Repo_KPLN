using KPLN_CoordiantorAI.Common;
using Microsoft.Win32;
using System;
using System.Windows;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class SettingsWindow : Window
    {
        private readonly CoordinatorAiRepository _repository;

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
            EmbeddingFilesTextBox.Text = settings.EmbeddingFilePaths;
            SystemPromptTextBox.Text = settings.SystemPrompt;
        }

        private void OnCreateDatabaseClick(object sender, RoutedEventArgs e)
        {
            try
            {
                _repository.EnsureDatabase();
                LoadGigaChatSettings();
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
                    EmbeddingFilePaths = EmbeddingFilesTextBox.Text,
                    SystemPrompt = SystemPromptTextBox.Text
                });

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

        private void OnBrowseEmbeddingFilesClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Выберите файлы эмбеддингов",
                CheckFileExists = true,
                Multiselect = true,
                Filter = "Файлы эмбеддингов (*.json;*.jsonl;*.txt;*.csv)|*.json;*.jsonl;*.txt;*.csv|Все файлы (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) == true)
                EmbeddingFilesTextBox.Text = string.Join(Environment.NewLine, dialog.FileNames);
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
    }
}