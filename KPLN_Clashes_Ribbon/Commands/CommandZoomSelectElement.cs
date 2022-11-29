using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Clashes_Ribbon.Commands
{
    public class CommandZoomSelectElement : IExecutableCommand
    {
        private int _id;

        private string _elInfo;
        
        public CommandZoomSelectElement(int id, string elInfo)
        {
            _id = id;
            _elInfo = elInfo;
        }

        
        public Result Execute(UIApplication app)
        {
            if (app.ActiveUIDocument != null)
            {
                Document doc = app.ActiveUIDocument.Document;
                Element element = doc.GetElement(new ElementId(_id));
                Transaction t = new Transaction(doc, "Zoom");
                
                if (!ElementCheckErrorFromInfoParse(doc, element))
                {
                    t.Start();
                    if (element != null)
                    {
                        ZoomTools.ZoomElement(element.get_BoundingBox(null), app.ActiveUIDocument);
                        app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId> { element.Id });
                    }
                    else
                        TaskDialog.Show("Внимание!", "Данный элемент не найден! Возможно, элемент был удален, или замоделирован заново, что привело к замене id.");
                    t.Commit();
                    
                    return Result.Succeeded;
                }
                
                return Result.Failed;
            }

            return Result.Failed;
        }

        /// <summary>
        /// Проверка элемента из отчета на вложенность в другой проект
        /// </summary>
        /// <returns>Да - элемент из связи; Нет - элемент из открытого проекта</returns>
        private bool ElementCheckErrorFromInfoParse(Document doc, Element element)
        {
            List<string> infosList = _elInfo.Split('➜').ToList();
            
            string nwcFileName = infosList.Where(i => i.Contains(".nwc")).FirstOrDefault();
            if (nwcFileName != null)
            {
                string rvtFileName = doc.PathName.Split(new string[] { ".rvt" }, StringSplitOptions.None)[0];
                nwcFileName = nwcFileName.Split(new string[] { ".nwc" }, StringSplitOptions.None)[0].Trim();
                if (!rvtFileName.Contains(nwcFileName))
                {
                    TaskDialog.Show("Внимание!", "Данный элемент находится в связи! Нельзя выбирать элементы из связи");
                    return true;
                }
            }

            int rvtLinksCount = infosList.Where(i => i.Contains(".rvt")).Count();
            if (rvtLinksCount > 0)
            {
                TaskDialog.Show("Внимание!", "Это элемент из связанного файла, который импортирован в ваш проект из другого ВАШЕГО файла" +
                    " (т.е. проект был разделен на части в Revit согласно требованиям ВЕР)");
                return true;
            }

            Category elCat = element.Category;
            if (elCat != null)
            {
                string catName = elCat.Name;
                if (catName.Equals(infosList[4].Trim()) | catName.Equals(infosList[5].Trim()))
                    return false;
                else
                    TaskDialog.Show("Внимание!", "Это элемент из связанного файла, но его id совпадает с id элемента в ВАШЕМ файле." +
                        "\nВид не будет соответвовать элементу из отчета");
            }
            
            return false;
        }
    }

}
