using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Tools.Common;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.Forms.Models.Core;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class RLinkManagerForm : Window
    {
        private static string _initialDirectoryForOpenFileDialog;

        private readonly UIApplication _uiapp;
        private readonly int _revitVersion;

        public RLinkManagerForm(UIApplication uiapp)
        {
            _uiapp = uiapp;
            Document doc = _uiapp.ActiveUIDocument.Document;
            _revitVersion = int.Parse(uiapp.Application.VersionNumber);

            ModelPath docModelPath = doc.GetWorksharingCentralModelPath() ?? throw new Exception("Работает только с моделями из хранилища");
            string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);

            DBProject dBProject = DBWorkerService.CurrentProjectDbService.GetDBProject_ByRevitDocFileName(strDocModelPath);
            if (dBProject != null) InitialDirectoryForOpenFileDialog = dBProject.MainPath;
            else InitialDirectoryForOpenFileDialog = Path.GetPathRoot(Environment.SystemDirectory);

            InitializeComponent();
            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Кэширование пути для выбора файлов с сервера
        /// </summary>
        public static string InitialDirectoryForOpenFileDialog 
        {
            get => _initialDirectoryForOpenFileDialog;
            set
            {
                // Очистка пути, если был передан файл
                if (new FileInfo(value).Exists)
                    _initialDirectoryForOpenFileDialog = new FileInfo(value).DirectoryName;
                else
                    _initialDirectoryForOpenFileDialog = value;
            }
        }

        /// <summary>
        /// Доп. настройки - UserControl
        /// </summary>
        public IRLinkUserControl SelectedConfig { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedConfig == null || SelectedConfig.LinkChangeEntityColl.Count() == 0)
                return;

            if (SelectedConfig.LinkChangeEntityColl.Any(ent => ent.CurrentEntStatus == EntityStatus.CriticalError || ent.CurrentEntStatus == EntityStatus.Error))
            {
                UserDialog ud = new UserDialog(
                    "ВНИМАНИЕ",
                    $"Операция отменена, т.к. в связях на замену есть ошибки (причины указаны в описаниях связей). Проверь, и устрани ошибки",
                    "Ошибки выделены оранжевым");
                ud.ShowDialog();

                return;
            }
            else if (SelectedConfig is RLinkUpdateContent && SelectedConfig.LinkChangeEntityColl.Any(ent => ent.CurrentEntStatus != EntityStatus.MarkedAsFinal))
            {
                UserDialog ud = new UserDialog(
                    "ВНИМАНИЕ",
                    $"Операция отменена, т.к. не все связи помечены как итоговый результат для замены. Проверь, и пометь связи как итоговый результат",
                    "Ошибки выделены оранжевым");
                ud.ShowDialog();

                return;
            }
            
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandLinkChanger_Start(SelectedConfig.LinkChangeEntityColl.ToArray()));
            Close();
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            MainScrollViewer.ScrollToVerticalOffset(MainScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void AddNewServLink_Click(object sender, RoutedEventArgs e)
        {
            ResetConfigToLoad();

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Revit Files (*.rvt)|*.rvt",
                InitialDirectory = InitialDirectoryForOpenFileDialog,
            };

            if (openFileDialog.ShowDialog() == true)
            {
                InitialDirectoryForOpenFileDialog = openFileDialog.FileName;
                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = Path.GetFileName(filePath);
                    LinkManagerLoadEntity newEntity = new LinkManagerLoadEntity(fileName, filePath);

                    SelectedConfig.AddNewItem(newEntity);
                }
            }

            DataContext = this;
        }

        private void AddNewRevitServerLink_Click(object sender, RoutedEventArgs e)
        {
            ResetConfigToLoad();

            ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(_revitVersion);
            if (rsFilesPickForm == null)
                return;

            if ((bool)rsFilesPickForm.ShowDialog())
            {
                foreach (ElementEntity formEntity in rsFilesPickForm.SelectedElements)
                {
                    LinkManagerLoadEntity newEntity = new LinkManagerLoadEntity(formEntity.Name,
                        $"RSN:\\\\{SelectFilesFromRevitServer.CurrentRevitServer.Host}{formEntity.Name}");

                    SelectedConfig.AddNewItem(newEntity);
                }

                DataContext = this;
            }
        }


        private void ChangeLink_Click(object sender, RoutedEventArgs e)
        {
            ResetConfigToUpdate();

            LinkManagerEntity[] docLMEnts = LoadRLI_Service.CreateLMEntities(_uiapp);
            foreach (LinkManagerEntity lmEnt in docLMEnts)
            {
                SelectedConfig.AddNewItem(lmEnt);
            }

            DataContext = this;
        }

        /// <summary>
        /// Обновляю и аннулирую если меняется класс SelectedConfig
        /// </summary>
        private void ResetConfigToLoad()
        {
            // Переключение и пересоздание основного конфига только при другом типе
            if (SelectedConfig == null || SelectedConfig is RLinkUpdateContent)
            {
                SelectedConfig = new RLinkLoadContent();
                DataContext = SelectedConfig;
            }
        }

        /// <summary>
        /// Обновляю и аннулирую если меняется класс SelectedConfig
        /// </summary>
        private void ResetConfigToUpdate()
        {
            // Переключение и пересоздание основного конфига как при другом типе, так и при существующем (обнуление)
            if (SelectedConfig == null || SelectedConfig is RLinkLoadContent || SelectedConfig is RLinkUpdateContent)
            {
                SelectedConfig = new RLinkUpdateContent(_revitVersion);
                DataContext = SelectedConfig;
            }
        }
    }
}
