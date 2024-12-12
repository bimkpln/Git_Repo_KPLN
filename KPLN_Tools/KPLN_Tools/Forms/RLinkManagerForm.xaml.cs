using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Tools.Common;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.ExecutableCommand;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class RLinkManagerForm : Window
    {
        private readonly UIApplication _uiapp;
        private readonly Document _doc;
        private readonly DBProject _dBProject;

        /// <summary>
        /// Кэширование пути для выбора файлов с сервера
        /// </summary>
        private string _initialDirectoryForOpenFileDialog = Path.GetPathRoot(Environment.SystemDirectory);

        private readonly int _revitVersion;
        private readonly bool _isLocalConfig = true;
        private readonly string _cofigName = "RLinkManager";

        public RLinkManagerForm(UIApplication uiapp)
        {
            _uiapp = uiapp;
            _doc = _uiapp.ActiveUIDocument.Document;
            _revitVersion = int.Parse(uiapp.Application.VersionNumber);

            ModelPath docModelPath = _doc.GetWorksharingCentralModelPath() ?? throw new Exception("Работает только с моделями из хранилища");
            string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);
            _dBProject = DBWorkerService.CurrentProjectDbService.GetDBProject_ByRevitDocFileName(strDocModelPath);

            if (_dBProject != null)
            {
                _initialDirectoryForOpenFileDialog = _dBProject.MainPath;
                _isLocalConfig = false;
            }

            LinkChangeEntityColl = new ObservableCollection<LinkManagerEntity>();

            InitializeComponent();
            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public ObservableCollection<LinkManagerEntity> LinkChangeEntityColl { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandLinkChanger_Start(LinkChangeEntityColl.ToArray()));
            Close();
        }

        private void SaveConfigBtn_Click(object sender, RoutedEventArgs e)
        {
            if (LinkChangeEntityColl.Count == 0)
            {
                MessageBox.Show("Нельзя сохранять пустую конфигурацию. Сначала - заполни её строками, а уже потом - сохраняй");
                return;
            }

            ConfigService.SaveConfig<LinkManagerEntity>(_doc, _cofigName, LinkChangeEntityColl, _isLocalConfig);

            MessageBox.Show("Конфигурации для проектов из этой папки сохранены успешно!");
        }

        private void LoadConfigBtn_Click(object sender, RoutedEventArgs e)
        {
            object obj = ConfigService.ReadConfigFile<ObservableCollection<LinkManagerEntity>>(_doc, _cofigName, _isLocalConfig);
            if (obj is IEnumerable<LinkManagerEntity> configItems && configItems.Any())
            {
                foreach (LinkManagerEntity item in configItems)
                {
                    LinkChangeEntityColl.Add(item);
                }
            }
            else
                MessageBox.Show("Конифгурация для проектов из этой папки еще не сохранялась");
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void AddNewServLink_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Revit Files (*.rvt)|*.rvt",
                InitialDirectory = _initialDirectoryForOpenFileDialog,
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _initialDirectoryForOpenFileDialog = openFileDialog.FileName;
                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = Path.GetFileName(filePath);
                    if (!LinkChangeEntityColl.Where(ent => ent.LinkPath == filePath).Any())
                        LinkChangeEntityColl.Add(new LinkManagerEntity(fileName, filePath));
                    else
                    {
                        CustomMessageBox cmb = new CustomMessageBox(
                            "Предупреждение",
                            $"Файл \"{fileName}\" по пути \"{filePath}\" уже присутсвует в конфигурации. Дублирование запрещено");
                        cmb.ShowDialog();
                    }
                }
            }
        }

        private void AddNewRevitServerLink_Click(object sender, RoutedEventArgs e)
        {
            ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(_revitVersion);
            if (rsFilesPickForm == null)
                return;

            bool? dialogResult = rsFilesPickForm.ShowDialog();
            if (dialogResult == null || rsFilesPickForm.Status != UIStatus.RunStatus.Run)
                return;

            foreach (ElementEntity formEntity in rsFilesPickForm.SelectedElements)
            {
                LinkChangeEntityColl.Add(new LinkManagerEntity(formEntity.Name, $"RSN:\\\\{SelectFilesFromRevitServer.CurrentRevitServer.Host}{formEntity.Name}"));
            }
        }

        private void ChangeLink_Click(object sender, RoutedEventArgs e)
        {

        }

        private void MenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is LinkManagerEntity linkChange)
                {
                    UserDialog ud = new UserDialog("ВНИМАНИЕ", $"Сейчас будут удалена загрузка файла \"{linkChange.LinkName}\". Продолжить?");
                    ud.ShowDialog();

                    if (ud.IsRun)
                        LinkChangeEntityColl.Remove(linkChange);
                }
            }
        }
    }
}
