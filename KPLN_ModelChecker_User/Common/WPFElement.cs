using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Common
{
    public class WPFElement
    {
        public Element Element { get; set; }
        
        /// <summary>
        /// Порядковый номер элемента в списке
        /// </summary>
        public string SerialNumber { get; set; }

        /// <summary>
        /// Имя элемента
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// Заголовок элемента
        /// </summary>
        public string Header { get; set; }
        
        /// <summary>
        /// Описание элемента
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// Текущий статус ошибки
        /// </summary>
        public Status CurrentStatus { get; set; }
        public string CurrentLevel { get; set; }
        public string QueryLevel { get; set; }
        public string QueryHeader { get; set; }
        public SolidColorBrush Fill { get; set; }
        public BoundingBoxXYZ Box { get; set; }
        public XYZ Centroid { get; set; }
        
        public WPFElement(Element element, string name, string header, string description, Status status)
        {
            Element = element;
            if (Element != null)
            {
                SerialNumber = element.Id.ToString();
            }
            else
            {
                SerialNumber = "#";
            }
            Name = name;
            Header = header;
            Description = description;

            CurrentStatus = status;
            switch (status)
            {
                case Status.AllmostOk:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 240, 240, 135));
                    Header = "Почти хорошо: " + header;
                    break;
                case Status.LittleWarning:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 220, 90));
                    Header = "Обрати внимание: " + header;
                    break;
                case Status.Warning:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 180, 90));
                    Header = "Предупреждение: " + header;
                    break;
                case Status.Error:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 125, 125));
                    Header = "Ошибка: " + header;
                    break;
            }
        }
        
        public WPFElement(Element element, string name, string header, string description, Status status, string currentLevel, string addValue, string addValueHeader, BoundingBoxXYZ box) : this (element, name, header, description, status)
        {
            try
            {
                Box = box;
                Centroid = new XYZ((box.Min.X + box.Max.X) / 2, (box.Min.Y + box.Max.Y) / 2, (box.Min.Z + box.Max.Z) / 2);
            }
            catch (Exception) { }
            CurrentLevel = currentLevel;
            QueryLevel = addValue;
            QueryHeader = addValueHeader;
        }
    }

}
