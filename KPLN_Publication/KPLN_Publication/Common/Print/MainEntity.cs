using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace KPLN_Publication
{
    /// <summary>
    /// Класс-"оболочка", хранящий лист Revit, сведения формате листа и параметры печати листа
    /// </summary>
    public class MainEntity : IComparable
    {
        /// <summary>
        /// Смылка на вид\лист
        /// </summary>
        public View MainView { get; private set; }

        public PaperSize RevitPaperSize { get; set; }
        
        public bool IsVertical { get; set; }
        
        public ElementId ViewId { get; private set; }
        
        public bool IsPrintable { get; set; }

        /// <summary>
        /// Параметры для печати нескольких листов на одном
        /// </summary>
        public List<FamilyInstance> TitleBlocks { get; set; }
        
        public double WidthMm { get; set; }
        
        public double HeigthMm { get; set; }

        public bool ForceColored { get; private set; }
        
        public string PdfFileName { get; set; }
        

        /// <summary>
        /// Инициализация класса, без объявления формата листа и параметров печати
        /// </summary>
        public MainEntity(View view)
        {
            MainView = view;
            ViewId = view.Id;

            ForceColored = false;
            Parameter isForceColoredParam = view.LookupParameter("Цветной");
            if(isForceColoredParam != null)
            {
                if(isForceColoredParam.HasValue)
                {
                    if (isForceColoredParam.AsInteger() == 1)
                        ForceColored = true;
                }
            }
        }

        public override string ToString()
        {
            if (MainView is ViewSheet sheet)
                return sheet.SheetNumber + " - " + sheet.Name;

            return MainView.Name;
        }

        /// <summary>
        /// Формирует имя листа на базе строки-"конструктора", содержащего имена параметров,
        /// которые будут заменены на значения параметров из данного листа
        /// </summary>
        /// <param name="constructor">Строка конструктора. Имена параметров должны быть включены в треугольные скобки.</param>
        /// <returns>Сформированное имя листа</returns>
        public string NameByConstructor(string constructor)
        {
            string name = "";

            string prefix = constructor.Split('<').First();
            name = name + prefix;

            string[] sa = constructor.Split('<');
            for(int i = 0; i < sa.Length; i++)
            {
                string s = sa[i];
                if (!s.Contains(">")) continue;

                string paramName = s.Split('>').First();
                string separator = s.Split('>').Last();

                string val = this.GetParameterValueBySheetOrProject(MainView, paramName);

                name = name + val;
                name = name + separator;
            }


            char[] arr = name
                .Where(c => 
                    (char.IsLetterOrDigit(c) 
                    || char.IsWhiteSpace(c) 
                    || c == '-' 
                    || c == '+'
                    || c == '_' 
                    || c == '.' 
                    || c == ','))
                .ToArray();

            name = new string(arr);

            return name;
        }

        /// <summary>
        /// Получает значение параметра из листа и из "информации о проекте", по аналогии с "меткой" в семействе основной надписи.
        /// </summary>
        /// <param name="sheet">Элемент модели</param>
        /// <param name="paramName">Имя параметра</param>
        /// <returns></returns>
        private string GetParameterValueBySheetOrProject(Element sheet, string paramName)
        {
            string value = "";

            Parameter param = sheet.LookupParameter(paramName);
            if(param == null)
            {
                param = sheet.Document.ProjectInformation.LookupParameter(paramName);
            }
            if(param != null)
            {
                value = this.GetParameterValueAsString(param);
            }
            return value;
        }

        /// <summary>
        /// Получает значение параметра с любым типом данных, преобразованное в тип string
        /// </summary>
        /// <param name="param">Имя параметра</param>
        /// <returns>Значение параметра как string</returns>
        private string GetParameterValueAsString(Parameter param)
        {
            string val = "";
            switch (param.StorageType)
            {
                case StorageType.None:
                    break;
                case StorageType.Integer:
                    val = param.AsInteger().ToString();
                    break;
                case StorageType.Double:
                    double d = param.AsDouble();
                    val = Math.Round(d * 304.8, 3).ToString();
                    break;
                case StorageType.String:
                    val = param.AsString();
                    break;
                case StorageType.ElementId:
                    val = param.AsElementId().ToString();
                    break;
                default:
                    break;
            }
            return val;
        }


        /// <summary>
        /// Попытаться преобразовать текстовый номер листа в число для правильной сортировки
        /// </summary>
        /// <returns></returns>
        public int GetSheetNumberAsInt(ViewSheet sheet)
        {
            string sheetNumberString = sheet.SheetNumber;

            if (sheetNumberString.Contains("-"))
            {
                sheetNumberString = sheetNumberString.Split('-').Last();
            }
            if(sheetNumberString.Contains("_"))
            {
                sheetNumberString = sheetNumberString.Split('_').Last();
            }
            int sheetNumber = 0;
            try
            {
                sheetNumber = Convert.ToInt32(System.Text.RegularExpressions.Regex.Replace(sheetNumberString, @"[^\d]+", ""));
            }
            catch
            {
            }
            return sheetNumber;
        }

        //для сортировки по номеру листа
        public int CompareTo(object obj)
        {
            MainEntity ms = obj as MainEntity;
            if(ms != null && ms.MainView is ViewSheet sheet)
            {
                int thisSheetNumber = this.GetSheetNumberAsInt(sheet);
                int compareSheetNumber = ms.GetSheetNumberAsInt(sheet);
                return thisSheetNumber.CompareTo(compareSheetNumber);
            }
            else if (ms != null && ms.MainView is View view)
            {
                string thisName = this.MainView.Name;
                string compareName = ms.MainView.Name;
                return thisName.CompareTo(compareName);
            }
            else
                throw new Exception("Невозможно сравнить два объекта");
        }
    }
}
