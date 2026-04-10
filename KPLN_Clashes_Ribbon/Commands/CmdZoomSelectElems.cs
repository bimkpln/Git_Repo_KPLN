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
    public class CmdZoomSelectElems<TId> : IExecutableCommand
        where TId : struct, IConvertible
    {
        private readonly TId[] _ids;
        private readonly string[] _elsInfo;

        public CmdZoomSelectElems(IEnumerable<TId> ids)
        {
            _ids = ids?.ToArray() ?? Array.Empty<TId>();
        }

        public CmdZoomSelectElems(TId id, string elInfo)
        {
            _ids = new[] { id };
            _elsInfo = new[] { elInfo ?? string.Empty };
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return Result.Cancelled;
            
            Document doc = app.ActiveUIDocument.Document;

            try
            {
                // Клик по одиночному ID - изоляция, зуммирование и выделение в модели (с пред проверкой)
                if (_ids.Length < 2 && _elsInfo.Length < 2)
                {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    Element elem = doc.GetElement(new ElementId(Convert.ToInt32(_ids.FirstOrDefault())));
#else
                    Element elem = doc.GetElement(new ElementId(Convert.ToInt64(_ids.FirstOrDefault())));
#endif
                    List<string> infosList = _elsInfo.FirstOrDefault().Split('➜').ToList();


                    Transaction t = new Transaction(doc, "KPLN_Приблизить");
                    if (!ElementCheckErrorFromInfoParse(doc, elem, infosList))
                    {
                        t.Start();
                        if (elem != null)
                        {
                            if (app.ActiveUIDocument.ActiveView is View3D activeView)
                                ZoomTools.ZoomElement(elem.get_BoundingBox(null), app.ActiveUIDocument, activeView);

                            SelectInDocTools.SelectElemsInDoc(uidoc, new List<ElementId> { elem.Id });
                        }
                        else
                            TaskDialog.Show("Внимание!", 
                                "Данный элемент не найден! Либо это элемент из связи, либо элемент был удален/замоделирован заново, что привело к удалению/замене id.\n\n" +
                                "ВАЖНО: Метку пересечения всё равно можно поставить, но ТОЛЬКО если она относится к открытому файлу и ТОЛЬКО на месте коллизии из отчёта");
                        t.Commit();
                    }
                }
                // Клик по коллекции ID - выделение в модели (без проверки на наличие)
                else
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    SelectInDocTools.SelectElemsInDoc(uidoc, _ids.Select(id => new ElementId(Convert.ToInt32(id))).ToList());
#else
                    SelectInDocTools.SelectElemsInDoc(uidoc, _ids.Select(id => new ElementId(Convert.ToInt64(id))).ToList());
#endif

                return Result.Succeeded;
                
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Cancelled;
            }
        }

        /// <summary>
        /// Проверка элемента из отчета на вложенность в другой проект
        /// </summary>
        /// <returns>Да - элемент из связи; Нет - элемент из открытого проекта</returns>
        private bool ElementCheckErrorFromInfoParse(Document doc, Element element, List<string> infosList)
        {
            if (element == null)
                return false;

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
