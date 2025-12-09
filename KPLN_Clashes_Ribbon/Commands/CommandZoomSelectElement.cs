using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Tools;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Clashes_Ribbon.Commands
{
    public class CommandZoomSelectElement : IExecutableCommand
    {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
        private readonly int _id;
#else
        private readonly long _id;
#endif

        private readonly string _elInfo;

        public CommandZoomSelectElement(int id, string elInfo)
        {
            _id = id;
            _elInfo = elInfo;
        }


        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            try
            {
                Document doc = app.ActiveUIDocument.Document;
                Element element = doc.GetElement(new ElementId(_id));
                Transaction t = new Transaction(doc, "KPLN_Приблизить");

                if (!ElementCheckErrorFromInfoParse(doc, element))
                {
                    t.Start();
                    if (element != null)
                    {
                        if (app.ActiveUIDocument.ActiveView is View3D activeView)
                            ZoomTools.ZoomElement(element.get_BoundingBox(null), app.ActiveUIDocument, activeView);

                        app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId> { element.Id });
                    }
                    else
                        TaskDialog.Show("Внимание!", 
                            "Данный элемент не найден! Либо это элемент из связи, либо элемент был удален/замоделирован заново, что привело к удалению/замене id.\n\n" +
                            "ВАЖНО: Метку пересечения всё равно можно поставить, но ТОЛЬКО если она относится к открытому файлу и ТОЛЬКО на месте коллизии из отчёта");
                    t.Commit();

                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Cancelled;
            }

            return Result.Cancelled;
        }

        /// <summary>
        /// Проверка элемента из отчета на вложенность в другой проект
        /// </summary>
        /// <returns>Да - элемент из связи; Нет - элемент из открытого проекта</returns>
        private bool ElementCheckErrorFromInfoParse(Document doc, Element element)
        {
            if (element == null)
                return false;

            List<string> infosList = _elInfo.Split('➜').ToList();

            // Тонкий отлов линков в отчете, при условии, что отчет идет в формате: xxx.nwc➜xxy.rvt (nwc всегда основной файл, rvt - это уже линк)
            string rvtLinkFileName = infosList.Where(i => i.ToLower().Contains(".rvt")).FirstOrDefault();
            if (rvtLinkFileName != null)
            {
                string rvtFileName = doc.PathName.ToLower().Split(new string[] { ".rvt" }, StringSplitOptions.None)[0];
                rvtLinkFileName = rvtLinkFileName.ToLower().Split(new string[] { ".rvt" }, StringSplitOptions.None)[0].Trim();
                if (!rvtFileName.Contains(rvtLinkFileName))
                {
                    TaskDialog.Show("Внимание!", "Данный элемент находится в связи! Нельзя выбирать элементы из связи");
                    return true;
                }
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

            TaskDialog.Show("Внимание!", "Элемент не прошел проверку. Скинь скрин отчета с отображением имени файла, " +
                "имени отчета и номером конфликта Куцко Тимофею. " +
                "Чтобы не тормозить работу - используй поиск по id (Управление ➜ Выбрать по коду)");
            return false;
        }
    }

}
