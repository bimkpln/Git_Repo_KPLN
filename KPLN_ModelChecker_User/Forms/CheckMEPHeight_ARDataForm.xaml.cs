using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class CheckMEPHeight_ARDataForm : Window
    {
        private readonly UIApplication _uiapp;
        private readonly Document _doc;
        private readonly string _configPath;

        public CheckMEPHeight_ARDataForm(
            UIApplication uiapp,
            CheckMEPHeightViewModel[] vModels)
        {
            _uiapp = uiapp;
            _doc = _uiapp.ActiveUIDocument.Document;
            ModelPath modelPath = _doc.GetWorksharingCentralModelPath() ?? throw new System.Exception("Работает только с моделями из хранилища");
            string folderPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath).Trim($"{_doc.Title}.rvt".ToArray());
            _configPath = folderPath + $"KPLNConfig\\IOS_CommandCheckMEPHeight.json";
            ViewModelColl = vModels;

            #region Заполняю поля окна в зависимости от наличия файла конфига
            // Файл конфига присутсвует. Читаю и чищу от неиспользуемых
            if (new FileInfo(_configPath).Exists)
            {
                CheckMEPHeightViewModel[] configVModels = ReadConfigFile();

                foreach (var confVM in configVModels)
                {
                    IEnumerable<CheckMEPHeightViewModel> matchingViewModels = ViewModelColl
                        .Where(vm =>
                            vm.VMCurrentRoomName.Equals(confVM.VMCurrentRoomName)
                            && vm.VMCurrentRoomDepartmentName.Equals(confVM.VMCurrentRoomDepartmentName));

                    foreach (var vm in matchingViewModels)
                    {
                        vm.VMIsCheckRun = confVM.VMIsCheckRun;
                        vm.VMCurrentRoomMinElemElevationForCheck = confVM.VMCurrentRoomMinElemElevationForCheck;
                        vm.VMCurrentRoomMinDistance = confVM.VMCurrentRoomMinDistance;
                    }
                }
            }
            #endregion


            InitializeComponent();
            ARDataItemContr.ItemsSource = ViewModelColl;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Коллекция-прослойка для группировки элементов в окне
        /// </summary>
        public CheckMEPHeightViewModel[] ViewModelColl { get; private set; }
        public bool IsRun { get; private set; } = false;

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                IsRun = false;
                Close();
            }
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            Close();
        }

        private void SaveConfigBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!new FileInfo(_configPath).Exists)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_configPath));
                FileStream fileStream = File.Create(_configPath);
                fileStream.Dispose();
            }

            if (ViewModelColl.Length > 0)
            {
                using (StreamWriter streamWriter = new StreamWriter(_configPath))
                {
                    string jsonEntity = JsonConvert.SerializeObject(ViewModelColl.Select(ent => ent.ToJson()));
                    streamWriter.Write(jsonEntity);
                }
            }

            MessageBox.Show("Сохаренено успешно!");
        }

        /// <summary>
        /// Десереилизация конфига
        /// </summary>
        private CheckMEPHeightViewModel[] ReadConfigFile()
        {
            List<CheckMEPHeightViewModel> entities = new List<CheckMEPHeightViewModel>();
            using (StreamReader streamReader = new StreamReader(_configPath))
            {
                string json = streamReader.ReadToEnd();
                entities = JsonConvert.DeserializeObject<List<CheckMEPHeightViewModel>>(json);
            }

            return entities.ToArray();
        }
    }
}
