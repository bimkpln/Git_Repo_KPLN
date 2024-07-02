using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using KPLN_Tools.Common.OVVK_System;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OVVK_PipeThocknessForm : Window
    {
        private readonly Document _doc;
        private readonly string _modelPath;
        private readonly string _configPath;

        public OVVK_PipeThocknessForm(
            Document document,
            Dictionary<ElementId, List<double>> pipeDict)
        {
            _doc = document;
            ModelPath modelPath = _doc.GetWorksharingCentralModelPath() ?? throw new System.Exception("Работает только с моделями из хранилища");
            _modelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath).Trim($"{_doc.Title}.rvt".ToArray());
            _configPath = _modelPath + $"OVVKConfig\\PipeThicknessConfig.json";

            #region Заполняю поля окна в зависимости от наличия файла конфига
            // Файл конфига присутсвует. Читаю и чищу от неиспользуемых
            if (new FileInfo(_configPath).Exists)
            {
                PipeThicknessEntities = ReadConfigFile();
                // Есть список имен типов труб, нужно назначить PipeType
                foreach (var entity in PipeThicknessEntities)
                {
                    foreach (var kvp in pipeDict)
                    {
                        if (_doc.GetElement(kvp.Key) is PipeType pipeType)
                        {
                            if (pipeType.Name == entity.CurrentPipeTypeName)
                            {
                                entity.CurrentPipeType = pipeType;
                                // Если нет такого диаметр в коллекции из проекта - PipeTypeDiamAndThickness нужно удалить
                                entity.CurrentPipeTypeDiamAndThickness.RemoveAll(thckn => !kvp.Value.Contains(thckn.CurrentDiameter));
                                
                                break;
                            }
                        }
                    }
                }

                // Если тип не назначен, или нет такого диаметр в коллекции из проекта - Entity нужно удалить
                PipeThicknessEntities.RemoveAll(ent => ent.CurrentPipeType == null);
            }
            
            // Уточняю конфиг по элементам, которые добавились в проект
            foreach (var kvp in pipeDict)
            {
                if(_doc.GetElement(kvp.Key) is PipeType pipeType)
                {
                    List<double> docDiams = kvp.Value;
                    PipeThicknessEntity exsitingFromConfigEntity = PipeThicknessEntities.FirstOrDefault(ent => ent.CurrentPipeTypeName == pipeType.Name);
                    
                    // Уточняю данные типов из конфига по типам из проекта
                    if (exsitingFromConfigEntity != null)
                    {
                        List<PipeTypeDiamAndThickness> exsitingFromConfigDiamAndThickness = exsitingFromConfigEntity.CurrentPipeTypeDiamAndThickness;

                        // Находим НЕсовпадающие диаметры (т.е. которых НЕТ в конфиге) и добавляем их
                        IEnumerable<double> differingDiams = docDiams
                            .Where(diam => !exsitingFromConfigDiamAndThickness.Any(exist => exist.CurrentDiameter == diam));
                        foreach (double diam in differingDiams)
                        {
                            exsitingFromConfigEntity.CurrentPipeTypeDiamAndThickness.Add(new PipeTypeDiamAndThickness(diam));
                        }
                    }
                    // Уточняю данными по типам из проекта, которых НЕ было в конфиге
                    else
                    {
                        PipeThicknessEntity entity = new PipeThicknessEntity(pipeType)
                        {
                            CurrentPipeTypeDiamAndThickness = docDiams.Select(diam => new PipeTypeDiamAndThickness(diam)).ToList()
                        };

                        PipeThicknessEntities.Add(entity);
                    }
                }
            }
            
            foreach (PipeThicknessEntity entity in PipeThicknessEntities)
            {
                entity.CurrentPipeTypeDiamAndThickness.Sort(CompareByDiameter);
            }

            #endregion

            InitializeComponent();
            PypeTypes.ItemsSource = PipeThicknessEntities;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public List<PipeThicknessEntity> PipeThicknessEntities { get; private set; } = new List<PipeThicknessEntity>();

        public bool IsRun { get; private set; } = false;

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                IsRun = false;
                Close();
            }
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
            
            if (PipeThicknessEntities.Count > 0)
            {
                using (StreamWriter streamWriter = new StreamWriter(_configPath))
                {
                    string jsonEntity = JsonConvert.SerializeObject(PipeThicknessEntities.Select(ent => ent.ToJson()));
                    streamWriter.Write(jsonEntity);
                }
            }

            MessageBox.Show("Сохаренено успешно!");
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        /// <summary>
        /// Десереилизация конфига
        /// </summary>
        private List<PipeThicknessEntity> ReadConfigFile()
        {
            List<PipeThicknessEntity> entities = new List<PipeThicknessEntity>();
            using (StreamReader streamReader = new StreamReader(_configPath))
            {
                string json = streamReader.ReadToEnd();
                entities = JsonConvert.DeserializeObject<List<PipeThicknessEntity>>(json);
            }

            return entities;
        }

        /// <summary>
        /// Проверяю текущую сущность на наличие аналога в проекте
        /// </summary>
        private bool CheckEntityForExsisting(List<double> docPipeDiams)
        {
            foreach (PipeThicknessEntity ent in PipeThicknessEntities)
            {
                foreach (PipeTypeDiamAndThickness thickness in ent.CurrentPipeTypeDiamAndThickness)
                {
                    if (docPipeDiams.Contains(thickness.CurrentDiameter))
                        return false;
                }
            }

            return true;
        }

        private int CompareByDiameter(PipeTypeDiamAndThickness entity1, PipeTypeDiamAndThickness entity2)
        {
            double diameter1 = entity1.CurrentDiameter;
            double diameter2 = entity2.CurrentDiameter;

            // Сравниваем CurrentDiameter двух объектов
            return diameter1.CompareTo(diameter2);
        }
    }
}
