using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Finishing.CommandTools;
using System;
using System.Collections.Generic;
using static KPLN_Finishing.Tools;

namespace KPLN_Finishing.ExternalCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    class Copy : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            if (Preferences.ShowCopyToolTip)
            {
                TaskDialog td = new TaskDialog("Привязать вручную");
                td.TitleAutoPrefix = false;
                td.MainContent = "Краткая инструкция";
                td.FooterText = "Алгоритмы работы:\n" +
                    "1) выбрать элемент(ы) которому нужно определить помещение - «Запуск скрипта» - выбрать помещение/элемент с рассчитанным помещением\n\n" +
                    "2) «Запуск скрипта» - выбрать элемент(ы) которому нужно определить помещение - выбрать помещение/элемент с рассчитанным помещением";
                td.CommonButtons = TaskDialogCommonButtons.Ok;
                td.VerificationText = "Больше не показывать";
                TaskDialogResult result = td.Show();
                if (td.WasVerificationChecked())
                { Preferences.ShowCopyToolTip = false; }
            }

            ICollection<ElementId> selectedElementsIds = commandData.Application.ActiveUIDocument.Selection.GetElementIds();
            if (selectedElementsIds.Count == 0)
            {
                try
                {
                    IList<Reference> refs = commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, new FilterNoRooms(), "Выберите нерассчитанные элементы <перекрытия> : <потолки> : <стены>");
                    foreach (Reference re in refs)
                    {
                        selectedElementsIds.Add(re.ElementId);
                    }
                }
                catch (Exception)
                {
                    return Result.Cancelled;
                }
            }
            Reference refer;
            try
            {
                refer = commandData.Application.ActiveUIDocument.Selection.PickObject(ObjectType.Element, new Filter(), "Выбрать рассчитанный элемент <перекрытия> : <потолки> : <стены> : <помещения>");
            }
            catch (Exception)
            {
                return Result.Cancelled;
            }
            Element selectedElement = doc.GetElement(refer.ElementId);
            using (Transaction t = new Transaction(doc, Names.assembly))
            {
                t.Start();
                foreach (ElementId id in selectedElementsIds)
                {
                    try
                    {
                        if (selectedElement.Category.Id.IntegerValue == -2000160)
                        {
                            Element element = doc.GetElement(id);
                            if (element.Category.Id.IntegerValue == -2000095)//if group
                            {
                                Group group = element as Group;
                                foreach (ElementId groupElementId in group.GetMemberIds())
                                {
                                    Element elementInGroup = doc.GetElement(groupElementId);
                                    ApplyRoom(elementInGroup, selectedElement);
                                }
                            }
                            else
                            {
                                if (element.Category.Id.IntegerValue == -2000011 || element.Category.Id.IntegerValue == -2000038 || element.Category.Id.IntegerValue == -2000032)
                                {
                                    ApplyRoom(element, selectedElement);
                                }
                            }
                        }
                        else
                        {
                            Element element = doc.GetElement(id);
                            if (element.Category.Id.IntegerValue == -2000095)//if group
                            {
                                Group group = element as Group;
                                foreach (ElementId groupElementId in group.GetMemberIds())
                                {
                                    Element elementInGroup = doc.GetElement(groupElementId);
                                    CopyFrom(doc, elementInGroup, selectedElement);
                                }
                            }
                            else
                            {
                                if (element.Category.Id.IntegerValue == -2000011 || element.Category.Id.IntegerValue == -2000038 || element.Category.Id.IntegerValue == -2000032)
                                {
                                    CopyFrom(doc, element, selectedElement);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    { }
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
}
