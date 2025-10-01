using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Forms.Entities;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ModelChecker_Lib.Forms
{
    /// <summary>
    /// Interaction logic for CheckMEPHeight_ARDataForm.xaml
    /// </summary>
    public partial class CheckMEPHeight_ARDataForm : Window
    {
        private readonly Document _doc;
        private readonly string _configPath;

        public CheckMEPHeight_ARDataForm(
            Document doc,
            CMHEntity[] vModels)
        {
            _doc = doc;
            ModelPath modelPath = _doc.GetWorksharingCentralModelPath() ?? throw new System.Exception("Работает только с моделями из хранилища");
            string folderPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath).Trim($"{_doc.Title}.rvt".ToArray());
            _configPath = folderPath + $"KPLN_Config\\IOS_CommandCheckMEPHeight.json";
            ViewModelColl = vModels;

            #region Заполняю поля окна в зависимости от наличия файла конфига
            // Файл конфига присутсвует. Читаю и чищу от неиспользуемых
            if (new FileInfo(_configPath).Exists)
            {
                CMHEntity[] configVModels = ReadConfigFile();

                foreach (var confVM in configVModels)
                {
                    IEnumerable<CMHEntity> matchingViewModels = ViewModelColl
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
        }

        /// <summary>
        /// Коллекция-прослойка для группировки элементов в окне
        /// </summary>
        public CMHEntity[] ViewModelColl { get; private set; }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
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
        private CMHEntity[] ReadConfigFile()
        {
            List<CMHEntity> entities = new List<CMHEntity>();
            using (StreamReader streamReader = new StreamReader(_configPath))
            {
                string json = streamReader.ReadToEnd();
                entities = JsonConvert.DeserializeObject<List<CMHEntity>>(json);
            }

            return entities.ToArray();
        }
    }
}
