using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Forms.Models
{
    public class OVVK_SystemManager_VM : INotifyPropertyChanged
    {
        /// <summary>
        /// Коллекция BuiltInCategory используемых в моделях ОВВК типа FamilyInstance
        /// </summary>
        public static BuiltInCategory[] FamileInstanceBICs = new BuiltInCategory[]
        {
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_Sprinklers,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_GenericModel,
        };

        /// <summary>
        /// Коллекция BuiltInCategory используемых в моделях ОВВК типа MEPCurve
        /// </summary>
        public static BuiltInCategory[] MEPCurveBICs = new BuiltInCategory[]
        {
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_PipeInsulations,
        };

        /// <summary>
        /// Функция для фильтрации семейств в коллекциях FilteredElementCollector
        /// </summary>
        public static Func<FamilyInstance, bool> FamilyNameFilter = x =>
            !x.Symbol.FamilyName.StartsWith("500_")
            && !x.Symbol.FamilyName.StartsWith("501_")
            && !x.Symbol.FamilyName.StartsWith("502_")
            && !x.Symbol.FamilyName.StartsWith("503_")
            && !x.Symbol.FamilyName.StartsWith("ASML_О_Отверстие_")
            && !x.Symbol.FamilyName.StartsWith("ClashPoint");

        public event PropertyChangedEventHandler PropertyChanged;

        private string _parameterName = "!!!<ВЫБЕРИ ПАРАМЕТР СИСТЕМЫ>!!!";
        private string _sysNameSeparator = "/";

        public OVVK_SystemManager_VM(Document doc)
        {
            CurrentDoc = doc;

            UpdateElementColl();

            SystemSumParameters = new ObservableCollection<string>()
            {
                "!!!<Выбери пар-р системы для предварительного просмотра списка систем проекта>!!!"
            };
        }

        /// <summary>
        /// Ревит-файл для окна
        /// </summary>
        public Document CurrentDoc { get; set; }

        /// <summary>
        /// Коллекция элементов для обработки в окне
        /// </summary>
        public List<Element> ElementColl { get; set; }

        /// <summary>
        /// Коллекция элементов с проблемами в назначении систем
        /// </summary>
        public List<Element> WarningsElementColl { get; set; } = new List<Element>();

        /// <summary>
        /// Имя параметра в окне
        /// </summary>
        public string ParameterName
        {
            get => _parameterName;
            set
            {
                if (_parameterName != value)
                {
                    _parameterName = value;
                    OnPropertyChanged();

                    UpdateSystemParamData();
                }
            }
        }

        /// <summary>
        /// Символ разделителя
        /// </summary>
        public string SysNameSeparator
        {
            get => _sysNameSeparator;
            set
            {
                if (_sysNameSeparator != value)
                {
                    _sysNameSeparator = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Коллекция групп систем по выбранному параметру
        /// </summary>
        public ObservableCollection<string> SystemSumParameters { get; set; }

        /// <summary>
        /// Обновить основную коллекцию эл-в для обработки
        /// </summary>
        public void UpdateElementColl()
        {
            ElementColl = new List<Element>();
            // Коллекция для анализа
            foreach (BuiltInCategory bic in FamileInstanceBICs)
            {
                ElementColl.AddRange(new FilteredElementCollector(CurrentDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(FamilyNameFilter));
            }

            foreach (BuiltInCategory bic in MEPCurveBICs)
            {
                ElementColl.AddRange(new FilteredElementCollector(CurrentDoc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType());
            }
        }

        /// <summary>
        /// Обновление списка систем в зависимости от пар-ра системы
        /// </summary>
        public void UpdateSystemParamData()
        {
            List<string> sysDataHeap = new List<string>();
            WarningsElementColl.Clear();

            string emptySysParamWarningMsg = "<ВНИМАНИЕ!!! Присутсвуют пустые значения системы. Нужен предварительный анализ, чтобы все было под контролем>";
            foreach (Element elem in ElementColl)
            {
                Parameter sysParam = elem.LookupParameter(_parameterName);
                if (sysParam == null)
                {
                    UserDialog ud = new UserDialog("ОШИБКА", $"Отсутсвует параметр {_parameterName} у эл-та: {elem.Id}.\n" +
                        $"Устрани ошибку, и повтори запуск. Анализ систем экстренно ЗАВЕРШЕН!");
                    ud.ShowDialog();
                    return;
                }

                string sysParamData = sysParam.AsString();
                // Анализ и корректировка значения пар-ра для вложенных общих семейств
                if (elem is FamilyInstance famInst && string.IsNullOrEmpty(sysParamData) && famInst.SuperComponent != null)
                {
                    Element parentElem = RecGetMainParentElement(famInst);
                    Parameter sysParentParam = parentElem.LookupParameter(_parameterName);
                    sysParamData = sysParentParam.AsString();
                }

                if (string.IsNullOrEmpty(sysParamData))
                {
                    sysDataHeap.Add(emptySysParamWarningMsg);
                    WarningsElementColl.Add(elem);
                }
                else if (!string.IsNullOrEmpty(sysParamData) && !sysDataHeap.Contains(sysParamData))
                    sysDataHeap.Add(sysParamData);
            }

            SystemSumParameters = new ObservableCollection<string>(sysDataHeap
                .GroupBy(x => x)
                .Select(x => x.Key)
                .OrderBy(p => p));

            OnPropertyChanged(nameof(SystemSumParameters));
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        /// <summary>
        /// Поиск основного родительского семейства для элемента
        /// </summary>
        private Element RecGetMainParentElement(FamilyInstance famInst)
        {
            if (famInst.SuperComponent != null)
            {
                Element supElem = famInst.SuperComponent;
                if (supElem is FamilyInstance fi)
                    return RecGetMainParentElement(fi);
                else
                    throw new System.Exception("Скинь разработчику - ОШИБКА приведения вложенного элемента");
            }
            else
                return famInst;
        }
    }
}
