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
        private List<SharedEntity> _errorList = new List<SharedEntity>();

        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров
        /// </summary>
        private string[] _separArr = new string[]
        {
            "...",
            "до",
            "-"
        };

        /// <summary>
        /// Точность размеров раздела АР
        /// </summary>
        private float _arAccuracy = 1.0f;
        /// <summary>
        /// Точность размеров раздела КР
        /// </summary>
        private float _krAccuracy = 0.1f;
        /// <summary>
        /// Точность размеров разделов ИОС
        /// </summary>
        private float _iosAccuracy = 1.0f;
        /// <summary>
        /// Точность межосевых размеров
        /// </summary>
        private float _gridsAccuracy = 0.000001f;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            string docTitle = doc.Title;

            //Основная часть
            FilteredElementCollector docDimensions = new FilteredElementCollector(doc).OfClass(typeof(Dimension)).WhereElementIsNotElementType();
            CheckOverride(docDimensions);
            CheckAccuracy(docDimensions, docTitle);

            ShowResult(doc);
            return Result.Succeeded;
        }


        /// <summary>
        /// Определяю размеры, которые были переопределены в проекте
        /// </summary>
        private void CheckOverride(FilteredElementCollector docDimensions)
        {
            foreach (Dimension dim in docDimensions)
            {
                // Игнорирую чертежные виды
                try
                {
                    if (dim.View.GetType().Equals(typeof(ViewDrafting))) { continue; }
                }
                catch (NullReferenceException) { continue; }

                double? currentValue = dim.Value;
                if (currentValue.HasValue && dim.ValueOverride?.Length > 0)
                {
                    double value = currentValue.Value * 304.8;
                    string overrideValue = dim.ValueOverride;
                    double[] doubleValues = DigitValue(overrideValue);
                    double minValue = doubleValues[0];
                    double maxValue = doubleValues[1];

                    // Нахожу значения без диапозона и игнорирую небольшие округления - больше 10 мм, при условии, что это не составляет 5% от размера
                    if (maxValue == Convert.ToDouble(int.MaxValue))
                    {
                        if (Math.Abs(minValue - value) > 10.0 || Math.Abs((minValue / value) * 100 - 100) > 5)
                        {
                            _errorList.Add(new SharedEntity(
                                dim.Id.IntegerValue,
                                dim.Name,
                                "Размер значительно отличается от реального",
                                $"Значение реального размера {Math.Round(value, 2)} мм, а при переопределении указано {overrideValue} мм. Разница существенная, лучше устранить.",
                                Status.Error)
                            );
                        }
                        else
                        {
                            _errorList.Add(new SharedEntity(
                                dim.Id.IntegerValue,
                                dim.Name,
                                "Размер не значительно отличается от реального",
                                $"Значение реального размера {Math.Round(value, 2)} мм, а при переопределении указано {overrideValue} мм. Разница не существенная, достаточно проконтролировать.",
                                Status.LittleWarning)
                            );
                        }
                    }
                    // Нахожу значения вне диапозоне
                    else if ((value >= maxValue | value < minValue))
                    {
                        _errorList.Add(new SharedEntity(
                            dim.Id.IntegerValue,
                            dim.Name,
                            "Размер вне диапозона",
                            $"Значение реального размера {Math.Round(value, 2)} мм, а диапозон указан {overrideValue}",
                            Status.Error)
                        );
                    }
                }
                else
                {
                    DimensionSegmentArray segments = dim.Segments;
                    foreach (DimensionSegment segment in segments)
                    {
                        if (segment.ValueOverride?.Length > 0)
                        {
                            _errorList.Add(new SharedEntity(
                                dim.Id.IntegerValue,
                                "Размер",
                                "Размер переопределен",
                                "Запрещено переопределять значения размеров",
                                Status.Error)
                            );
                        }
                    }
                }

            }
        }

        /// <summary>
        /// Поиск цифровых значений
        /// </summary>
        /// <param name="value">Текст, который нужно распознать</param>
        /// <returns>Массив из 2-х значений: 1 - Минимальное значение, а также текущее для бездиапозонных; 2 - максимальное значение (в случае наличия дапозонов)</returns>
        private double[] DigitValue(string value)
        {
            string ovverrideMinValue = String.Empty;
            double overrideMinDouble = 0.0;
            string ovverrideMaxValue = String.Empty;
            int overrideMaxInt = int.MaxValue;
            
            string[] splitValues = value.Split(_separArr, StringSplitOptions.None); 
            if (splitValues.Length > 1)
            {
                ovverrideMinValue = splitValues[0];
                ovverrideMaxValue = splitValues[1];
            }
            else
            {
                ovverrideMinValue = value;
            }

            string onlyNumbMin = new string(ovverrideMinValue.Where(x => Char.IsDigit(x)).ToArray());
            overrideMinDouble = Double.Parse(onlyNumbMin);
            
            if (ovverrideMaxValue.Equals(String.Empty))
            {
                return new double[2] { overrideMinDouble, Convert.ToDouble(overrideMaxInt) };
            }
            else
            {
                string onlyNumbMax = new string(ovverrideMaxValue.Where(x => Char.IsDigit(x)).ToArray());
                double overrideMaxDouble = Double.Parse(onlyNumbMax);
                return new double[2] { overrideMinDouble, overrideMaxDouble};
            }
        }

        /// <summary>
        /// Пользовательский срез по разделителю
        /// </summary>
        /// <param name="value">Текст, который нужно распознать</param>
        /// <param name="separ">Разделитель</param>
        /// <returns></returns>
        private string[] MyRegex(string value, string separ)
        {
            string pattern = $@"\b[{separ}]+";
            string[] trySplit = Regex.Split(value, pattern);
            return trySplit;
        }
        
        /// <summary>
        /// Определяю размеры, которые не соответствуют необходимой точности
        /// </summary>
        private void CheckAccuracy(FilteredElementCollector docDimensions, string docTitle)
        {
            foreach (Dimension dim in docDimensions)
            {
                if (_errorList.Where(x => x.Id == dim.Id.IntegerValue).ToList().Count > 0)
                {
                    continue;
                }
                
                double? currentValue = dim.Value;
                if (currentValue.HasValue)
                {
                    DimensionType dimType = dim.DimensionType;
                    FormatOptions dimTypeSettings = dimType.GetUnitsFormatOptions();
                    //double currentAccuracy = dimTypeSettings.Accuracy;
                    //!!! ДОБИТЬ ОТСЮДА. БЫЛО ПРИНЯТО РЕШЕНИЕ - ИСКАТЬ ТИПОРАЗМЕРЫ РАЗМЕРОВ, И ОПРЕДЕЛЯТЬ ИХ ОКРУГЛЕНИЕ. ДАЛЕЕ, ЕСЛИ ОКРУГЛЕНИЕ ДО 1 (Т.Е. В НОРМЕ), ТО ДЛЯ КР 
                    // НУЖНО ПРОИЗВЕСТИ СРАВНЕНИЕ РЕАЛЬНОГО ЗНАЧЕНИЯ, С ЦЕЛЬЮ ВЫЯВЛЕНИЯ ЕГО ТОЧНОСТИ
                    // ПОФИКСИТЬ ТОТ ФАКТ, ЧТО ДЛЯ ЭЛЕМЕНТОВ НА ЛЕГЕНДЕ - АВТОМАТИЧЕСКИЙ ВИД НЕ ОТКРЫВАЕТСЯ

                }
                else
                {
                    DimensionSegmentArray segments = dim.Segments;
                    foreach (DimensionSegment segment in segments)
                    {
                        
                    }
                }

            }
        }

        /// <summary>
        /// Метод для вывода результатов пользователю
        /// </summary>
        private void ShowResult(Document doc)
        {
            if (_errorList.Count == 0)
            {
                TaskDialog.Show("Результат", "Проблемы не обнаружены :)", TaskDialogCommonButtons.Ok);
            }
            else
            {
                // Настраиваю сортировку в окне и генерирую экземпляры ошибок
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                ObservableCollection<WPFDisplayItem> wpfFiltration = new ObservableCollection<WPFDisplayItem>();
                wpfFiltration.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });
                foreach (SharedEntity entity in _errorList)
                {
                    Element element = doc.GetElement(new ElementId(entity.Id));
                    WPFDisplayItem item = GetItemByElement(element, entity.Name, element.OwnerViewId.IntegerValue, entity.Header, entity.Description, entity.CurrentStatus);
                    item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Id элемента: ", Description = entity.Id.ToString() });
                    outputCollection.Add(item);
                    // Добавляю критерии сортировки
                    //wpfFiltration.Add(new WPFDisplayItem(-2, StatusExtended.Critical, kvp.Key.Id.IntegerValue) { Name = $"Лист номер {kvp.Key.SheetNumber}" });
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
            WPFDisplayItem item = new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus, elemId);
            try
            {
                item.SetZoomParams(element, null);
                item.Name = name;
                item.Header = header;
                item.Description = description;
                item.Category = string.Format("<{0}>", element.Category.Name);
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Подсказка: ", Description = description });
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
    }
}
