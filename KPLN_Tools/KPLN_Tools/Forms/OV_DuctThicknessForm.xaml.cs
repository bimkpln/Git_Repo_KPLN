using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.ExecutableCommand;
using Newtonsoft.Json;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OV_DuctThicknessForm : Window
    {
        private readonly Document _doc;
        private readonly Element[] _elementsToSet;
        private readonly string _configPath;

        public OV_DuctThicknessForm(Document doc, Element[] elementsToSet)
        {
            _doc = doc;
            _elementsToSet = elementsToSet;
            
            ModelPath docModelPath = doc.GetWorksharingCentralModelPath() ?? throw new System.Exception("Работает только с моделями из хранилища");
            string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath).Trim($"{doc.Title}.rvt".ToArray());
            _configPath = strDocModelPath + $"KPLN_Config\\OV_DuctThickness.json";

            #region Заполняю поля окна в зависимости от наличия файла конфига
            // Файл конфига присутсвует
            if (new FileInfo(_configPath).Exists)
                ReadConfigFile();
            else
            {
                CurrentDuctThicknessEntity = new DuctThicknessEntity()
                {
                    ParameterName = "КП_И_Толщина стенки",
                    PartOfInsulationName = "EI",
                    PartOfSystemName = "_Противодым._",
                };
            }
            #endregion

            InitializeComponent();

            this.DataContext = CurrentDuctThicknessEntity;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public DuctThicknessEntity CurrentDuctThicknessEntity { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            if (e.Key == Key.Enter)
                StartBtn_Click(sender, e);
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            Task.Run(() => { SaveConfig(); });

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandDuctThickness_Start(CurrentDuctThicknessEntity, _elementsToSet));

            Close();
        }

        private void BtnParamSearch_Click(object sender, RoutedEventArgs e)
        {
            ElementSinglePick paramForm = SelectParameterFromRevit.CreateForm(_doc, _elementsToSet, StorageType.Double);
            paramForm.ShowDialog();

            if (paramForm.SelectedElement != null)
                CurrentDuctThicknessEntity.ParameterName = paramForm.SelectedElement.Name;
        }

        /// <summary>
        /// Сериализация и сохранение файла-конфигурации
        /// </summary>
        private void SaveConfig()
        {
            if (!new FileInfo(_configPath).Exists)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_configPath));
                FileStream fileStream = File.Create(_configPath);
                fileStream.Dispose();
            }

            using (StreamWriter streamWriter = new StreamWriter(_configPath))
            {
                object entObj = CurrentDuctThicknessEntity.ToJson();
                string jsonEntity = JsonConvert.SerializeObject(entObj, Formatting.Indented);
                streamWriter.Write(jsonEntity);
            }
        }

        /// <summary>
        /// Десереилизация конфига
        /// </summary>
        private void ReadConfigFile()
        {
            using (StreamReader streamReader = new StreamReader(_configPath))
            {
                string json = streamReader.ReadToEnd();
                CurrentDuctThicknessEntity = JsonConvert.DeserializeObject<DuctThicknessEntity>(json);
            }
        }
    }
}
