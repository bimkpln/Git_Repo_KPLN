using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Forms.Models
{
    public class OVVK_SystemManager_ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _parameterName = "!!!<ВЫБЕРИ ПАРАМЕТР СИСТЕМЫ>!!!";
        private string _sysNameSeparator = "/";

        public OVVK_SystemManager_ViewModel(Document doc, List<Element> elems)
        {
            CurrentDoc = doc;
            ElementColl = elems;

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

                    UpdateSystemSumParams(_parameterName);
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

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        /// <summary>
        /// Обновление списка систем в зависимости от пар-ра системы
        /// </summary>
        private void UpdateSystemSumParams(string paramName)
        {
            List<string> sysDataHeap = new List<string>();
            WarningsElementColl.Clear();

            string emptySysParamWarningMsg = "<ВНИМАНИЕ!!! Присутсвуют пустые значения системы. Нужен предварительный анализ, чтобы все было под контролем>";
            foreach (Element elem in ElementColl)
            {
                Parameter sysParam = elem.LookupParameter(paramName);
                if (sysParam == null)
                {
                    UserDialog ud = new UserDialog("ОШИБКА", $"Отсутсвует параметр {paramName} у эл-та: {elem.Id}.\n" +
                        $"Устрани ошибку, и повтори запуск. Анализ систем экстренно ЗАВЕРШЕН!");
                    ud.ShowDialog();
                    return;
                }

                string sysParamData = sysParam.AsString();
                // Анализ и корректировка значения пар-ра для вложенных общих семейств
                if (elem is FamilyInstance famInst && string.IsNullOrEmpty(sysParamData) && famInst.SuperComponent != null)
                {
                    Element parentElem = RecGetMainParentElement(famInst);
                    Parameter sysParentParam = parentElem.LookupParameter(paramName);
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
