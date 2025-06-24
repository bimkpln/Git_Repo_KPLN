using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using KPLN_Library_ConfigWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Tools.Common;
using KPLN_Tools.Common.OVVK_System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OVVK_PipeThicknessForm : Window
    {
        private readonly Document _doc;
        private readonly DBProject _dBProject;

        private readonly ConfigType _configType = ConfigType.Local;
        private readonly string _cofigName = "OVVK_PipeThicknessConfig";

        public OVVK_PipeThicknessForm(
            Document document,
            Dictionary<ElementId, List<double>> pipeDict)
        {
            _doc = document;

            ModelPath docModelPath = _doc.GetWorksharingCentralModelPath() ?? throw new Exception("Работает только с моделями из хранилища");
            string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);
            _dBProject = DBWorkerService.CurrentProjectDbService.GetDBProject_ByRevitDocFileName(strDocModelPath);

            if (_dBProject != null)
                _configType = ConfigType.Shared;

            #region Заполняю поля окна в зависимости от наличия файла конфига
            // Файл конфига присутсвует
            object obj = ConfigService.ReadConfigFile<List<PipeThicknessEntity>>(_doc, _configType, _cofigName);
            if (obj is IEnumerable<PipeThicknessEntity> configItems && configItems.Any())
            {
                PipeThicknessEntities = configItems.ToList();

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
                if (_doc.GetElement(kvp.Key) is PipeType pipeType)
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
            if (PipeThicknessEntities.Count == 0)
            {
                MessageBox.Show("Нельзя сохранять пустую конфигурацию. Сначала - заполни её строками, а уже потом - сохраняй");
                return;
            }

            ConfigService.SaveConfig<PipeThicknessEntity>(_doc, _configType, PipeThicknessEntities, _cofigName);

            MessageBox.Show("Конфигурации для проектов из этой папки сохранены успешно!");
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
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
