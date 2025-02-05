using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_Finishing.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Finishing.Tools;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Finishing.ExternalCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    class Update : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            if (new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements().Count == 0)
            {
                Print("В проекте отсутствуют помещения!\n\nПодсказка: Необходимо запускать плагин  в проектах с присутствующими помещениями и элементами отделки. Связанные файлы в расчете не принимаюит участия.", MessageType.Error);
                return Result.Cancelled;
            }
            TaskDialog td = new TaskDialog("Обновить ведомости");
            td.TitleAutoPrefix = false;
            td.MainContent = "Запустить менеджер обновления спецификации отделки?";
            td.FooterText = Names.task_dialog_hint;
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.Close;
            td.VerificationText = "Очистить существующие значения";
            TaskDialogResult result = td.Show();
            if (result != TaskDialogResult.Yes)
            { return Result.Cancelled; }
            if (!AllParametersExist(new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements().First(), doc))
            {
                Print("Не все параметры, требуемые для работы, подгружены в проект!\n\nПодсказка: Воспользуйтесь загрузчиком параметров либо проверьте параметры вручную.", MessageType.Error);
                return Result.Failed;
            }
            if (td.WasVerificationChecked())
            {
                int n1 = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements().Count() * 10;
                string s1 = "{0} из " + n1.ToString() + " элементов обработано";
                using (ProgressFormSimple pf = new ProgressFormSimple("Обнуление данных", s1, n1))
                {
                    using (Transaction t = new Transaction(doc, "KPLN Очистить значения"))
                    {
                        t.Start();
                        foreach (Element room in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements())
                        {
                            pf.SetInfoStrip(string.Format("Помещение: {0}", room.Id.ToString()));
                            foreach (string param in new string[] { "О_ПОМ_Ведомость",
                                                                    "О_ПОМ_ГОСТ_Описание стен",
                                                                    "О_ПОМ_ГОСТ_Описание стен_Текст",
                                                                    "О_ПОМ_ГОСТ_Описание полов",
                                                                    "О_ПОМ_ГОСТ_Описание полов_Текст",
                                                                    "О_ПОМ_ГОСТ_Описание потолков",
                                                                    "О_ПОМ_ГОСТ_Описание потолков_Текст",
                                                                    "О_ПОМ_ГОСТ_Описание плинтусов",
                                                                    "О_ПОМ_ГОСТ_Длина плинтусов_Текст",
                                                                    "О_ПОМ_Группа"})
                            {
                                pf.Increment();
                                try
                                {
                                    room.LookupParameter(param).Set("");
                                }
                                catch (Exception) { }
                            }
                        }
                        t.Commit();
                    }
                }
            }
            List<WPFParameter> wpfParameters = new List<WPFParameter>
            {
                new WPFParameter()
            };

            BuiltInParameter[] roomBIPs = new BuiltInParameter[] 
            {
                BuiltInParameter.ROOM_NAME,
                BuiltInParameter.EDITED_BY,
                BuiltInParameter.ROOM_PHASE,
                BuiltInParameter.LEVEL_NAME,
                BuiltInParameter.ROOM_NUMBER,
                BuiltInParameter.ROOM_DEPARTMENT,
                BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS 
            };

            Element[] docRooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .ToArray();

            foreach (BuiltInParameter builtInParameter in roomBIPs)
            {
                Parameter parameter = docRooms.FirstOrDefault().get_Parameter(builtInParameter);
                wpfParameters.Add(new WPFBuiltInParameter(builtInParameter, parameter.Definition.Name, parameter.StorageType));
            }

            if (doc.IsWorkshared)
            {
                BuiltInParameter wsBIP = BuiltInParameter.ELEM_PARTITION_PARAM;
                Parameter parameter = docRooms.FirstOrDefault().get_Parameter(wsBIP);
                wpfParameters.Add(new WPFBuiltInParameter(wsBIP, parameter.Definition.Name, parameter.StorageType));
            }

            foreach (WPFParameter parameter in GetLocalParameters(doc))
            {
                wpfParameters.Add(parameter);
            }
            List<MetaRoom> rooms = new List<MetaRoom>();
            List<MetaRoom> nonEmptyRooms = new List<MetaRoom>();
            List<MetaRoom> roomCollection = new List<MetaRoom>();
            foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements())
            {
                try
                {
                    Room room = element as Room;
                    if (room.Area > 0.0000001)
                    {
                        roomCollection.Add(new MetaRoom(element as Room));
                    }
                }
                catch (Exception) { }
            }
            int n = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements().Count;
            n += new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements().Count;
            n += new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements().Count;
            string s = "{0} из " + n.ToString() + " элементов обработано";
            using (ProgressFormSimple pf = new ProgressFormSimple("Подготовка элементов", s, n))
            {
                foreach (BuiltInCategory category in new BuiltInCategory[] { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Ceilings })
                {
                    foreach (Element element in new FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements())
                    {
                        pf.Increment();
                        if (GetGroupParameter(doc, element) != "отделка") { continue; }
                        try
                        {
                            MetaElement meta = new MetaElement(element);
                            int linkedRoomId = int.Parse(element.LookupParameter("О_Id помещения").AsString(), System.Globalization.NumberStyles.Integer);
                            MetaRoom room = roomCollection.Find(o => o.Id == linkedRoomId);
                            if (room != null)
                            {
                                room.AddElement(meta);
                            }
                        }
                        catch (Exception) { }
                    }
                }
            }
            foreach (MetaRoom room in roomCollection)
            {
                if (room.Elements.Count != 0)
                {
                    nonEmptyRooms.Add(room);
                }
            }
            if (nonEmptyRooms.Count != 0)
            {
                UpdateSetup form = new UpdateSetup(wpfParameters, nonEmptyRooms);
                form.Show();
                return Result.Succeeded;
            }
            else
            {
                Print("Не найдено ни одного помещения в проекте со связанными элементами отделки", MessageType.Error);
                return Result.Cancelled;
            }
        }
    }
}