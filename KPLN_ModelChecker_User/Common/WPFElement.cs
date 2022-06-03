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
        public string Id { get; set; }
        public string Name { get; set; }
        public string Header { get; set; }
        public string Description { get; set; }
        public string CurrentLevel { get; set; }
        public string QueryLevel { get; set; }
        public string QueryHeader { get; set; }
        public Element Element { get; set; }
        public SolidColorBrush Fill { get; set; }
        public BoundingBoxXYZ Box { get; set; }
        public XYZ Centroid { get; set; }
        public WPFElement(Element element, string name, string header, string description, string currentLevel, string addValue, string addValueHeader, Status status, BoundingBoxXYZ box)
        {
            Element = element;
            try
            {
                Box = box;
                Centroid = new XYZ((box.Min.X + box.Max.X) / 2, (box.Min.Y + box.Max.Y) / 2, (box.Min.Z + box.Max.Z) / 2);
            }
            catch (Exception) { }
            switch (status)
            {
                case Status.AllmostOk:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 240, 240, 135));
                    break;
                case Status.LittleWarning:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 220, 90));
                    break;
                case Status.Warning:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 180, 90));
                    break;
                case Status.Error:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 125, 125));
                    break;
                case Status.Ok:
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 75, 245, 170));
                    break;
            }
            if (Element != null)
            {
                Id = element.Id.ToString();
            }
            else
            {
                Id = "#";
            }
            Name = name;
            Header = header;
            Description = description;
            CurrentLevel = currentLevel;
            QueryLevel = addValue;
            QueryHeader = addValueHeader;
        }
    }
}
