using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools.Common.LinkManager;
using KPLN_Tools.ExecutableCommand;
using Microsoft.Win32;
using Newtonsoft.Json;
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
        /// <summary>
        /// Кэширование пути для выбора файлов с сервера
        /// </summary>
        private static string _initialDirectoryForOpenFileDialog = @"Y:\";
        private readonly UIApplication _uiapp;
        private readonly Document _doc;
        private readonly string _configPath;
        private int _revitVersion;

        public RLinkManagerForm(UIApplication uiapp)
        {
            _uiapp = uiapp;
            _doc = _uiapp.ActiveUIDocument.Document;
            _revitVersion = int.Parse(uiapp.Application.VersionNumber);

            string docPath;
            if (_doc.IsWorkshared)
                docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(_doc.GetWorksharingCentralModelPath());
            else
                docPath = _doc.PathName;

            string folderPath = docPath.Trim($"{_doc.Title}.rvt".ToArray());
            if (!Directory.Exists(folderPath))
                folderPath = @"C:\temp\";
            _initialDirectoryForOpenFileDialog = folderPath;
            _configPath = folderPath + $"KPLN_Config\\RLinkManager.json";
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

            if (!new FileInfo(_configPath).Exists)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_configPath));
                FileStream fileStream = File.Create(_configPath);
                fileStream.Dispose();
            }

            if (LinkChangeEntityColl.Count > 0)
            {
                using (StreamWriter streamWriter = new StreamWriter(_configPath))
                {
                    string jsonEntity = JsonConvert.SerializeObject(LinkChangeEntityColl.Select(ent => ent.ToJson()), Formatting.Indented);

                    streamWriter.Write(jsonEntity);
                }
            }

            MessageBox.Show("Конфигурации для проектов из этой папки сохранены успешно!");
        }

        private void LoadConfigBtn_Click(object sender, RoutedEventArgs e)
        {
            // Файл конфига присутсвует. Читаю и чищу от неиспользуемых
            if (new FileInfo(_configPath).Exists)
            {
                List<LinkManagerEntity> newEntities = ReadConfigFile();
                if (newEntities == null || newEntities.Count() == 0)
                {
                    MessageBox.Show("Файл-конфигурации пуст. Пересохрани его с заполненными строками, и повтори попытку");
                    return;
                }

                foreach (LinkManagerEntity entity in newEntities)
                {
                    if (!LinkChangeEntityColl.Any(ent => ent.LinkName.Equals(entity.LinkName)))
                        LinkChangeEntityColl.Add(entity);
                }
            }
            else
                MessageBox.Show("Конифгурация для проектов из этой папки еще не сохранялась");
        }

        /// <summary>
        /// Десереилизация конфига
        /// </summary>
        private List<LinkManagerEntity> ReadConfigFile()
        {
            List<LinkManagerEntity> entities = new List<LinkManagerEntity>();
            using (StreamReader streamReader = new StreamReader(_configPath))
            {
                string json = streamReader.ReadToEnd();

                entities = JsonConvert.DeserializeObject<List<LinkManagerEntity>>(json);
            }

            return entities;
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
