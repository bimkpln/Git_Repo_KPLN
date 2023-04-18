using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckDimensions : IExternalCommand
    {
        /// <summary>
        /// Список сущностей со своими атрибутами, которые соответсвуют ошибке
        /// </summary>
        private List<WPFElement> _errorList = new List<WPFElement>();

        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров
        /// </summary>
        private string[] _separArr = new string[]
        {
            "...",
            "до",
            "-",
            "max",
            "min"
        };

        ///// <summary>
        ///// Точность размеров раздела АР
        ///// </summary>
        //private float _arAccuracy = 1.0f;
        ///// <summary>
        ///// Точность размеров раздела КР
        ///// </summary>
        //private float _krAccuracy = 0.1f;
        ///// <summary>
        ///// Точность размеров разделов ИОС
        ///// </summary>
        //private float _iosAccuracy = 1.0f;
        ///// <summary>
        ///// Точность межосевых размеров
        ///// </summary>
        //private float _gridsAccuracy = 0.000001f;

        private WPFDisplayItem GetItemByElement(Element element, string name, int elemId, string header, string description, Status status)
        {
            StatusExtended exstatus;
            switch (status)
            {
                case Status.Error:
                    exstatus = StatusExtended.Critical;
                    break;
                case Status.LittleWarning:
                    exstatus = StatusExtended.LittleWarning;
                    break;
                default:
                    exstatus = StatusExtended.Warning;
                    break;
            }
            WPFDisplayItem item = new WPFDisplayItem(-20000260, exstatus, elemId);
            try
            {
                item.SetZoomParams(element, null);
                item.Name = name;
                item.Header = header;
                item.Description = description;
                item.Category = "Размеры";
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                item.Collection.Add(new WPFDisplayItem(-20000260, exstatus) { Header = "Подсказка: ", Description = description });
                HashSet<string> values = new HashSet<string>();
            }
            catch (Exception e)
            {
                try
                {
                    PrintError(e.InnerException);
                }
                catch (Exception) { }
                PrintError(e);
            }
            return item;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            
            //Основная часть
            FilteredElementCollector docDimensions = new FilteredElementCollector(doc).OfClass(typeof(Dimension)).WhereElementIsNotElementType();
            try
            {
                CheckOverride(doc, docDimensions);
                CheckAccuracy(doc);
                ShowResult();
            }
            catch (Exception ex)
            {
                PrintError(ex);
                return Result.Failed;
            }

            return Result.Succeeded;
        }


        /// <summary>
        /// Определяю размеры, которые были переопределены в проекте
        /// </summary>
        private void CheckOverride(Document doc, FilteredElementCollector docDimensions)
        {
            foreach (Dimension dim in docDimensions)
            {
                // Игнорирую чертежные виды
                if (dim.View == null)
                    continue;

                if (dim.View.GetType().Equals(typeof(ViewDrafting)))
                    continue;

                double? currentValue = dim.Value;
                
                if (currentValue.HasValue && dim.ValueOverride?.Length > 0)
                {
                    double value = currentValue.Value * 304.8;
                    string overrideValue = dim.ValueOverride;
                    int dimId = dim.Id.IntegerValue;
                    string dimName = dim.Name;
                    CheckDimValues(doc, value, overrideValue, dimId, dimName);
                }
                else
                {
                    DimensionSegmentArray segments = dim.Segments;
                    foreach (DimensionSegment segment in segments)
                    {
                        if (segment.ValueOverride?.Length > 0)
                        {
                            double value = segment.Value.Value * 304.8;
                            string overrideValue = segment.ValueOverride;
                            int dimId = dim.Id.IntegerValue;
                            string dimName = dim.Name;
                            CheckDimValues(doc, value, overrideValue, dimId, dimName);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Анализ значения размера
        /// </summary>
        private void CheckDimValues(Document doc, double value, string overrideValue, int elemId, string elemName)
        {

            string ovverrideMinValue = String.Empty;
            double overrideMinDouble = 0.0;
            string ovverrideMaxValue = String.Empty;
            double overrideMaxDouble = 0.0;

            string[] splitValues = overrideValue.Split(_separArr, StringSplitOptions.None);

            // Анализирую диапозоны
            if (splitValues.Length > 1)
            {
                ovverrideMinValue = splitValues[0];
                if (ovverrideMinValue.Length == 0)
                {
                    ovverrideMinValue = overrideValue;
                }
                else
                {
                    ovverrideMaxValue = splitValues[1];
                }

                string onlyNumbMin = new string(ovverrideMinValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out overrideMinDouble);
                if (!ovverrideMaxValue.Equals(String.Empty))
                {
                    string onlyNumbMax = new string(ovverrideMaxValue.Where(x => Char.IsDigit(x)).ToArray());
                    Double.TryParse(onlyNumbMax, out overrideMaxDouble);
                    // Нахожу значения вне диапозоне
                    if (value >= overrideMaxDouble | value < overrideMinDouble)
                    {
                        _errorList.Add(new WPFElement(
                            doc.GetElement(new ElementId(elemId)),
                            elemName,
                            "Размер вне диапозона",
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"",
                            Status.Error)
                        );
                    }
                }
                else
                {
                    _errorList.Add(new WPFElement(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Диапазон. Нужен ручной анализ",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"",
                        Status.Warning)
                    );
                }
            }

            // Нахожу значения без диапозона и игнорирую небольшие округления - больше 10 мм, при условии, что это не составляет 5% от размера
            else
            {
                double overrideDouble = 0.0;
                string onlyNumbMin = new string(overrideValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out overrideDouble);
                if (overrideDouble == 0.0)
                {
                    _errorList.Add(new WPFElement(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Нет возможности анализа",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\". Оцени вручную",
                        Status.Error)
                    );
                }
                else if (Math.Abs(overrideDouble - value) > 10.0 || Math.Abs((overrideDouble / value) * 100 - 100) > 5)
                {
                    _errorList.Add(new WPFElement(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Размер значительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница существенная, лучше устранить.",
                        Status.Error)
                    );
                }
                else
                {
                    _errorList.Add(new WPFElement(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Размер незначительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница не существенная, достаточно проконтролировать.",
                        Status.LittleWarning)
                    );
                }
            }
        }
        
        /// <summary>
        /// Определяю размеры, которые не соответствуют необходимой точности
        /// </summary>
        private void CheckAccuracy(Document doc)
        {
            string docTitle = doc.Title;
            FilteredElementCollector docDimensionTypes = new FilteredElementCollector(doc).OfClass(typeof(DimensionType)).WhereElementIsElementType();
            foreach (DimensionType dimType in docDimensionTypes)
            {
                if (dimType.UnitType == UnitType.UT_Length)
                {
                    FormatOptions typeOpt = dimType.GetUnitsFormatOptions();
                    try
                    {
                        double currentAccuracy = typeOpt.Accuracy;
                        if (currentAccuracy > 1.0)
                        {
                            _errorList.Add(new WPFElement(
                                doc.GetElement(new ElementId(dimType.Id.IntegerValue)),
                                dimType.Name,
                                "Размер имеет запрещенно низкую точность",
                                $"Принятое округление в 1 мм, а в данном типе - указано \"{currentAccuracy}\" мм. Замени округление, или удали типоразмер.",
                                Status.Error)
                            );
                        }
                    }
                    catch (Exception) 
                    { 
                        //Игнорирую типы без настроек округления
                    }
                }
            }
        }

        /// <summary>
        /// Метод для вывода результатов пользователю
        /// </summary>
        private void ShowResult()
        {
            if (_errorList.Count == 0)
            {
                Print("[Проверка размеров] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
            }
            else
            {
                // Настраиваю сортировку в окне и генерирую экземпляры ошибок
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                ObservableCollection<WPFDisplayItem> wpfFiltration = new ObservableCollection<WPFDisplayItem>();
                wpfFiltration.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });
                foreach (WPFElement wpfElem in _errorList)
                {
                    Element element = wpfElem.Element;
                    WPFDisplayItem item = GetItemByElement(wpfElem.Element, wpfElem.Name, element.OwnerViewId.IntegerValue, wpfElem.Header, wpfElem.Description, wpfElem.CurrentStatus);
                    item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Id элемента: ", Description = wpfElem.Element.Id.ToString() });
                    outputCollection.Add(item);
                }
                List<WPFDisplayItem> sortedOutputCollection = outputCollection.OrderBy(o => o.Header).ToList();

                // Вывожу результат
                int counter = 1;
                ObservableCollection<WPFDisplayItem> wpfElements = new ObservableCollection<WPFDisplayItem>();
                foreach (WPFDisplayItem e in sortedOutputCollection)
                {
                    e.Header = string.Format("{0}# {1}", (counter++).ToString(), e.Header);
                    wpfElements.Add(e);
                }
                if (wpfElements.Count != 0)
                {
                    ElementsOutputExtended form = new ElementsOutputExtended(wpfElements, wpfFiltration);
                    form.Show();
                }
            }
        }

    }
}
