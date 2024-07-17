using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Parameters_Ribbon.Common.CopyElemParamData;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using static System.Windows.Forms.AxHost;

namespace KPLN_Parameters_Ribbon.Command
{
    public class CommandWriteValues : IExecutableCommand
    {
        private readonly List<ParameterRuleElement> _rules;
        private readonly bool _isWriteInGroups;

        public CommandWriteValues(ObservableCollection<ParameterRuleElement> rules, bool isWriteInGroups)
        {
            _rules = rules.ToList();
            _isWriteInGroups = isWriteInGroups;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            int max = 0;
            foreach (ParameterRuleElement rule in _rules)
            {
                max += new FilteredElementCollector(doc)
                    .OfCategoryId((rule.SelectedCategory.Data as Category).Id)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .Count;
            }
            string format = "{0} из " + max.ToString() + " параметров обработано";
            Progress_Single pb = new Progress_Single("KPLN: Копирование параметров", format, true);
            pb.SetProggresValues(max, 0);
            pb.ShowProgress();

            try
            {
                #region Сеттинг параметров, для которых НЕ возможен вариант разных значений между группами, И они в группе
                if (_isWriteInGroups)
                {
                    #region Архивная наработка, которую не пустить в релиз - работает только с группами, которые не имеют пространственных разночтений с исходниокм (вращение)
                    // Проблема НЕ решается (ручные правки в модели, или попытка эмитации в ревит не приносят результатов - группы вращаются в положение исходной)
                    // Проблема должна быть решена расшиернием API, но все никак: https://forums.autodesk.com/t5/revit-ideas/group-edit-mode-in-api/idc-p/10293127

                    //using (Transaction t = new Transaction(doc, "KPLN: Коп пар-в: анализ групп"))
                    //{
                    //    t.Start();

                    //    GroupType[] groupTypeArr = new FilteredElementCollector(doc)
                    //        .OfClass(typeof(GroupType))
                    //        .Cast<GroupType>()
                    //        .ToArray();
                    //    HashSet<ElementId> delGroupsSet = new HashSet<ElementId>();
                    //    foreach (GroupType groupType in groupTypeArr)
                    //    {
                    //        GroupSet groupSet = groupType.Groups;
                    //        if (groupSet.Size != 0)
                    //        {
                    //            IEnumerator groupEnmu = groupSet.GetEnumerator();
                    //            // Получаю новую группу
                    //            int count = 0;
                    //            while (groupEnmu.MoveNext())
                    //            {
                    //                Group currentNewGroup = null;
                    //                count++;

                    //                if (groupEnmu.Current is Group currentGroup)
                    //                {
                    //                    string currentGroupOldName = currentGroup.Name;
                    //                    ElementId[] groupElemIDsArr = currentGroup.UngroupMembers().ToArray();
                    //                    foreach (ParameterRuleElement rule in _rules)
                    //                    {
                    //                        Category selectedRuleCat = rule.SelectedCategory.Data as Category ?? throw new Exception($"Не удалось получить категрию элемнтов для одного из правил. Обратись к разработчику");

                    //                        foreach (ElementId elemId in groupElemIDsArr)
                    //                        {
                    //                            Element elem = doc.GetElement(elemId);

                    //                            if (elem is SketchPlane || elem is SketchBase)
                    //                                continue;

                    //                            Category elemCat = elem.Category;
                    //                            if (elemCat == null)
                    //                                Print($"Не удалось получить категрию элемнта с ID: {elemId} и именем {elem.Name}. Он пропущен в анализе, нужен ручной анализ со стороны пользователя", MessageType.Error);
                    //                            else if (elemCat.Id == selectedRuleCat.Id)
                    //                            {
                    //                                Parameter sourceParameter = GetParameterByElement(elem, rule.SelectedSourceParameter);
                    //                                Parameter targetParameter = GetParameterByElement(elem, rule.SelectedTargetParameter);

                    //                                if (targetParameter.Definition is InternalDefinition intParamDef)
                    //                                {
                    //                                    // Сеттинг параметров, для которых НЕ возможен вариант разных значений между группами (иной вариант устанавливается ниже)
                    //                                    if (!intParamDef.VariesAcrossGroups)
                    //                                    {
                    //                                        SetElemParamByRule(sourceParameter, targetParameter, pb);
                    //                                        pb.AddProgress(groupSet.Size - 1);
                    //                                        delGroupsSet.Add(currentGroup.Id);
                    //                                    }
                    //                                }
                    //                            }
                    //                        }
                    //                    }

                    //                    string currentGroupNewName = $"{currentGroupOldName}_ZhvBlr{count}";
                    //                    currentNewGroup = doc.Create.NewGroup(groupElemIDsArr);
                    //                    currentNewGroup.GroupType.Name = currentGroupNewName;
                    //                }
                    //            }

                    //        }
                    //    }

                    //    // Удаляю старую версию
                    //    foreach(ElementId groupId in delGroupsSet)
                    //    {
                    //        doc.Delete(groupId);
                    //    }

                    //    t.Commit();
                    //}

                    //using (Transaction t = new Transaction(doc, "KPLN: Коп пар-в: очистка групп"))
                    //{
                    //    t.Start();

                    //    GroupType[] groupTypeArr = new FilteredElementCollector(doc)
                    //        .OfClass(typeof(GroupType))
                    //        .Cast<GroupType>()
                    //        .ToArray();

                    //    foreach (GroupType groupType in groupTypeArr)
                    //    {
                    //        GroupType tempGroupType = groupType;
                    //        string[] splitedNameArr = groupType.Name.Split(new string[] { "_ZhvBlr" }, StringSplitOptions.None);
                    //        if (splitedNameArr.Length > 1)
                    //        {
                    //            string cleareadNewGroupName = splitedNameArr[0];
                    //            GroupType[] groupClearedTypeNames = new FilteredElementCollector(doc)
                    //                .OfClass(typeof(GroupType))
                    //                .Cast<GroupType>()
                    //                .Where(gt => gt.Name == cleareadNewGroupName)
                    //                .ToArray();

                    //            GroupSet groupSet = groupType.Groups;
                    //            if (groupSet.Size != 0)
                    //            {
                    //                IEnumerator groupEnmu = groupSet.GetEnumerator();
                    //                // Получаю новую группу
                    //                bool isDeletable = false;
                    //                while (groupEnmu.MoveNext())
                    //                {
                    //                    if (groupEnmu.Current is Group currentGroup)
                    //                    {
                    //                        if (!groupClearedTypeNames.Any())
                    //                            groupType.Name = cleareadNewGroupName;
                    //                        else
                    //                        {
                    //                            tempGroupType = groupClearedTypeNames.FirstOrDefault();
                    //                            isDeletable = true;
                    //                        }

                    //                        LocationPoint oldLocation = currentGroup.Location as LocationPoint;
                    //                        Line oldAxis = Line.CreateUnbound(oldLocation.Point, XYZ.BasisZ);
                    //                        currentGroup.GroupType = tempGroupType;
                    //                        doc.Regenerate();

                    //                        LocationPoint newLocation = currentGroup.Location as LocationPoint;
                    //                        Line newAxis = Line.CreateUnbound(newLocation.Point, XYZ.BasisZ);

                    //                        XYZ translation = oldLocation.Point - newLocation.Point;
                    //                        ElementTransformUtils.MoveElement(doc, currentGroup.Id, translation);


                    //                    }
                    //                }

                    //                // Удаляю старую версию
                    //                if (isDeletable)
                    //                    doc.Delete(groupType.Id);
                    //            }

                    //        }
                    //    }

                    //    t.Commit();
                    //}
                    #endregion

                    using (Transaction t = new Transaction(doc, "KPLN: Коп пар-в: ВСЕ"))
                    {
                        t.Start();

                        foreach (ParameterRuleElement rule in _rules)
                        {
                            Element[] elemArr = new FilteredElementCollector(doc)
                                .OfCategoryId((rule.SelectedCategory.Data as Category).Id)
                                .WhereElementIsNotElementType()
                                .ToArray();

                            foreach (Element element in elemArr)
                            {
                                Parameter sourceParameter = GetParameterByElement(element, rule.SelectedSourceParameter);
                                Parameter targetParameter = GetParameterByElement(element, rule.SelectedTargetParameter);
                                if (targetParameter.Definition is InternalDefinition intParamDef)
                                    SetElemParamByRule(sourceParameter, targetParameter, pb);
                                else
                                    throw new Exception($"Не удалось преобразовать в InternalDefinition для парамтера {targetParameter.Definition.Name}");
                            }
                        }

                        t.Commit();
                    }

                }
                #endregion

                #region Сеттинг параметров, для которых возможен вариант разных значений между группами, или параметр может отличаться между экземплярами групп 
                using (Transaction t = new Transaction(doc, "KPLN: Коп пар-в: ВНЕ групп"))
                {
                    t.Start();

                    foreach (ParameterRuleElement rule in _rules)
                    {
                        Element[] elemArr = new FilteredElementCollector(doc)
                            .OfCategoryId((rule.SelectedCategory.Data as Category).Id)
                            .WhereElementIsNotElementType()
                            .ToArray();

                        foreach (Element element in elemArr)
                        {
                            Parameter sourceParameter = GetParameterByElement(element, rule.SelectedSourceParameter);
                            Parameter targetParameter = GetParameterByElement(element, rule.SelectedTargetParameter);
                            if (targetParameter.Definition is InternalDefinition intParamDef)
                            {
                                // Сеттинг параметров, для которых возможен вариант разных значений между группами
                                if (intParamDef.VariesAcrossGroups)
                                    SetElemParamByRule(sourceParameter, targetParameter, pb);
                                // Сеттинг параметров, для которых НЕ возможен вариант разных значений между группами, НО они не в группе
                                else if (element.GroupId.IntegerValue == -1)
                                    SetElemParamByRule(sourceParameter, targetParameter, pb);
                            }
                            else
                                throw new Exception($"Не удалось преобразовать в InternalDefinition для парамтера {targetParameter.Definition.Name}");
                        }

                    }

                    t.Commit();
                }
                #endregion
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                pb.Update(0, "ОТКОЛНЕНО");

                PrintError(ex);

                return Result.Cancelled;
            }
            finally
            {
                pb.SetBtn_Ok_Enabled();
            }
        }

        private static Parameter GetParameterByElement(Element element, ListBoxElement rule)
        {
            foreach (Parameter p in element.Parameters)
            {
                if (p.Definition.Name == (rule.Data as Parameter).Definition.Name)
                {
                    return p;
                }
            }

            Element elemType = element.Document.GetElement(element.GetTypeId());
            if (elemType != null)
            {
                foreach (Parameter p in elemType.Parameters)
                {
                    if (p.Definition.Name == (rule.Data as Parameter).Definition.Name)
                    {
                        return p;
                    }
                }
            }

            return null;
        }

        private static void SetElemParamByRule(Parameter sourceParameter, Parameter targetParameter, Progress_Single pb)
        {
            pb.Increment("Заполнение параметров");

            if (sourceParameter != null && targetParameter != null)
            {
                switch (targetParameter.StorageType)
                {
                    case StorageType.Double:
                        double? dv = GetDoubleValue(sourceParameter);
                        if (dv != null)
                        {
                            targetParameter.Set((double)dv);
                        }
                        break;
                    case StorageType.Integer:
                        int? iv = GetIntegerValue(sourceParameter);
                        if (iv != null)
                        {
                            targetParameter.Set((int)iv);
                        }
                        break;
                    case StorageType.String:
                        string sv = GetStringValue(sourceParameter);
                        if (sv != null && sv != " " && sv != string.Empty)
                        {
                            targetParameter.Set(sv);
                        }
                        break;
                }
            }
        }

        private static double? GetDoubleValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        return p.AsDouble();
                    case StorageType.Integer:
                        return p.AsInteger();
                    case StorageType.String:
                        return double.Parse(p.AsString(), System.Globalization.NumberStyles.Float);
                    default:
                        return null;
                }
            }
            catch (Exception) { return null; }
        }

        private static int? GetIntegerValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        return (int)Math.Round(p.AsDouble());
                    case StorageType.Integer:
                        return p.AsInteger();
                    case StorageType.String:
                        return int.Parse(p.AsString(), System.Globalization.NumberStyles.Integer);
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string GetStringValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        return p.AsValueString();
                    case StorageType.Integer:
                        return p.AsValueString();
                    case StorageType.String:
                        return p.AsString();
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                try
                {
                    switch (p.StorageType)
                    {
                        case StorageType.Double:
                            return p.AsDouble().ToString();
                        case StorageType.Integer:
                            return p.AsInteger().ToString();
                        case StorageType.String:
                            return p.AsString();
                        default:
                            return null;
                    }
                }
                catch (Exception) { return null; }
            }
        }
    }
}
