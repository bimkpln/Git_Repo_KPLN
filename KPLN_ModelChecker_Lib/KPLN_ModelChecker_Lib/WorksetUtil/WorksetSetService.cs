using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.WorksetUtil.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_ModelChecker_Lib.WorksetUtil
{
    public class WorksetSetService
    {
        /// <summary>
        /// Глобальный путь к папке с конфигами
        /// </summary>
        public static string WSPatternFolderPath;

        /// <summary>
        /// Сервис по созданию рабочих наборов в проекте
        /// </summary>
        /// <param name="doc">Проект для создания РН</param>
        /// <param name="rvtLinks">Список связей, для которых нужно создать РНы</param>
        /// <param name="modelElemWSCreating">Метка создания РН для элементов модели (кроме экземпляров связей)</param>
        public static bool ExecuteFromService(
            Document doc,
            IEnumerable<RevitLinkInstance> rvtLinks,
            IEnumerable<DirectShape> dirShapes,
            IEnumerable<PointCloudInstance> pcInstances,
            IEnumerable<ImportInstance> importInstances,
            bool impInstWSCreating = true,
            bool modelElemWSCreating = true)
        {
            // Кастомная настройка пути под проекты
            if (doc.Title.StartsWith("СЕТ_1"))
                WSPatternFolderPath = $"X:\\BIM\\5_Scripts\\Git_Repo_KPLN\\KPLN_ModelChecker_Debugger\\KPLN_ModelChecker_Debugger\\Workset_Patterns\\СЕТ";
            else
                WSPatternFolderPath = $"X:\\BIM\\5_Scripts\\Git_Repo_KPLN\\KPLN_ModelChecker_Debugger\\KPLN_ModelChecker_Debugger\\Workset_Patterns";

            #region Вывод пользовательского окна с xml-шаблонами
            System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog
            {
                InitialDirectory = WSPatternFolderPath,
                Multiselect = false,
                Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*"
            };
            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return false;

            string xmlFilePath = dialog.FileName;

            //Десериализация пользовательского xml-файла
            WorksetDTO dto = new WorksetDTO();
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(WorksetDTO));
            using (StreamReader r = new StreamReader(xmlFilePath))
            {
                dto = (WorksetDTO)serializer.Deserialize(r);
            }
            #endregion

            #region Первичное создание рабочих наборов в модели
            if (!doc.IsWorkshared && doc.CanEnableWorksharing())
            {
                TaskDialog taskDialog = new TaskDialog("Внимание!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainContent = "В данном документе не включена функция совместной работы. Хотите активировать функцию совместной работы и создать рабочие наборы?",
                    CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
                };
                TaskDialogResult dialogResult = taskDialog.Show();
                if (dialogResult == TaskDialogResult.Cancel)
                    return false;
                else
                {
                    // Попытка поиска РН для уровней
                    string abrDepartment = "аббрРазд";
                    string gridsLevelsWS = "аббрРазд_Оси и уровни";
                    WorksetByCurrentParameter currentGridsLevelsWS = dto.WorksetByCurrentParameterList
                        .Where(p => p.WorksetName.Contains("Оси и уровни"))
                        .FirstOrDefault();
                    if (currentGridsLevelsWS != null)
                    {
                        gridsLevelsWS = currentGridsLevelsWS.WorksetName;
                        abrDepartment = gridsLevelsWS.Split('_')[0];
                    }

                    // Попытка поиска дефолтного РН для раздела
                    string defaultWS;
                    WorksetByCurrentParameter currentDefaultWS = dto.WorksetByCurrentParameterList
                        .Where(p => !p.WorksetName.Contains("Оси и уровни") && p.WorksetName.Contains(abrDepartment))
                        .FirstOrDefault();
                    if (currentDefaultWS != null)
                        defaultWS = currentDefaultWS.WorksetName;
                    else
                        defaultWS = $"{abrDepartment}_Модель";

                    // Создаю первичные РН в проекте
                    doc.EnableWorksharing(gridsLevelsWS, defaultWS);
                }
            }
            #endregion

            # region Создание рабочих наборов по конфигам
            using (Transaction t = new Transaction(doc))
            {
                t.Start("KPLN_Создание рабочих наборов");

                // Общие поля класса WorksetDTO
                string linkedFilesPrefix = dto.LinkedFilesPrefix;
                string dwgLinksName = dto.DWGLinksName;
                bool useMonitoredElems = dto.UseMonitoredElements;

                //Назначение рабочих наборов для связанных файлов
                foreach (RevitLinkInstance linkInstance in rvtLinks)
                {
                    if (linkInstance.IsValidObject && doc.GetElement(linkInstance.GetTypeId()) is RevitLinkType linkFileType)
                    {
                        if (linkFileType.IsNestedLink) continue;

                        string linkWorksetName1 = linkInstance.Name.Split(':')[0];
                        string linkWorksetName2 = linkWorksetName1.Substring(0, linkWorksetName1.Length - 5);
                        string linkWorksetName = linkedFilesPrefix + linkWorksetName2;
                        Workset linkWorkset = CreateNewWorkset(doc, linkWorksetName);

                        WorksetByCurrentParameter.SetWorkset(linkInstance, linkWorkset);
                        WorksetByCurrentParameter.SetWorkset(linkFileType, linkWorkset);
                    }
                }

                //Назначение рабочих наборов для координационных моделей
                foreach (DirectShape dirShape in dirShapes)
                {
                    string linkWorksetName = linkedFilesPrefix + "КоордМодель_" + dirShape.Name.Split(new string[] { ".nw" }, StringSplitOptions.None)[0];
                    Workset linkWorkset = CreateNewWorkset(doc, linkWorksetName);

                    WorksetByCurrentParameter.SetWorkset(dirShape, linkWorkset);
                }

                //Назначение рабочих наборов для облака точек
                foreach (PointCloudInstance pcInstance in pcInstances)
                {
                    string linkWorksetName = linkedFilesPrefix + "ОблТочек_" + pcInstance.Name.Split(new string[] { ".rcs" }, StringSplitOptions.None)[0];
                    Workset linkWorkset = CreateNewWorkset(doc, linkWorksetName);

                    WorksetByCurrentParameter.SetWorkset(pcInstance, linkWorkset);
                }

                //Назначение рабочих наборов для dwg-импорта/связей
                if (impInstWSCreating)
                {
                    foreach (ImportInstance importInstance in importInstances)
                    {
                        // Обход конфиогов, где НЕ нужно генерить РН под cad (например: Сетунь)
                        if (string.IsNullOrEmpty(dwgLinksName))
                            break;

                        if (doc.GetElement(importInstance.GetTypeId()) is CADLinkType cadLinkType)
                        {
                            Workset cadWorkset = CreateNewWorkset(doc, dwgLinksName);

                            WorksetByCurrentParameter.SetWorkset(importInstance, cadWorkset);
                            WorksetByCurrentParameter.SetWorkset(cadLinkType, cadWorkset);
                        }
                    }
                }

                // Назначение рабочих наборов по WorksetByCurrentParameter
                foreach (WorksetByCurrentParameter param in dto.WorksetByCurrentParameterList)
                {
                    Workset workset = param.GetWorkset(doc);

                    //Назначение рабочих наборов по категории
                    if (param.BuiltInCategories.Count != 0 && modelElemWSCreating)
                    {
                        foreach (BuiltInCategory bic in param.BuiltInCategories)
                        {
                            List<Element> elems = new FilteredElementCollector(doc)
                                .OfCategory(bic)
                                .WhereElementIsNotElementType()
                                .ToElements()
                                .ToList();

                            foreach (Element elem in elems)
                            {
                                WorksetByCurrentParameter.SetWorkset(elem, workset);
                            }
                        }
                    }

                    //Назначение рабочих наборов по имени семейства
                    if (param.FamilyNames.Count != 0 && modelElemWSCreating)
                    {
                        List<FamilyInstance> famIns = new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .ToList();
                        foreach (string familyName in param.FamilyNames)
                        {
                            List<FamilyInstance> elems = famIns
                                .Where(f => f.Symbol.FamilyName.ToLower().Contains(familyName.ToLower()))
                                .ToList();

                            foreach (Element elem in elems)
                            {
                                WorksetByCurrentParameter.SetWorkset(elem, workset);
                            }
                        }
                    }

                    //Назначение рабочих наборов по имени типа
                    if (param.TypeNames.Count != 0 && modelElemWSCreating)
                    {
                        List<Element> allModelElements = new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .Cast<Element>()
                            .ToList();
                        foreach (string typeName in param.TypeNames)
                        {
                            foreach (Element elem in allModelElements)
                            {
                                ElementId typeId = elem.GetTypeId();
                                if (typeId == null || typeId == ElementId.InvalidElementId)
                                    continue;

                                if (!(doc.GetElement(typeId) is ElementType elemType))
                                    continue;

                                if (elemType.Name.ToLower().Contains(typeName.ToLower()))
                                {
                                    WorksetByCurrentParameter.SetWorkset(elem, workset);
                                }
                            }
                        }
                    }

                    //Назначение рабочих наборов по заполненному параметру
                    if (param.SelectedParameters.Count != 0 && modelElemWSCreating)
                    {
                        List<Element> allModelElements = new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .Cast<Element>()
                            .ToList();

                        foreach (SelectedParameter p in param.SelectedParameters)
                        {
                            foreach (Element elem in allModelElements)
                            {
                                try
                                {
                                    string data;
                                    try
                                    {
                                        data = elem.LookupParameter(p.ParameterName).AsString();
                                    }
                                    catch (ArgumentNullException)
                                    {
                                        data = elem.LookupParameter(p.ParameterName).AsValueString();
                                    }
                                    if (data.Equals(p.ParameterValue))
                                    {
                                        WorksetByCurrentParameter.SetWorkset(elem, workset);
                                    }
                                }
                                catch (NullReferenceException) { }
                            }
                        }

                    }
                }

                t.Commit();
            }
            #endregion

            // ВОЗМОЖНО УДАЛЕНИЕ РН ТОЛЬКО НАЧИНАЯ С РЕВИТ2024, до этого - нет API
            //EmptyWorksetsForm emptyWorksetsForm = new EmptyWorksetsForm(doc);
            //if (emptyWorksetsForm.EmptyWorksets.Any())
            //    emptyWorksetsForm.Show();

            List<string> emptyWorksetsNames = GetEmptyWorksets(doc);
            if (emptyWorksetsNames.Count > 0)
            {
                string msg = "Обнаружены пустые рабочие наборы! Их следует удалить вручную:\n";
                foreach (string s in emptyWorksetsNames)
                {
                    msg += s + "\n";
                }
                TaskDialog.Show("Отчёт", msg);
            }

            return true;
        }

        /// <summary>
        /// Метод для поиска и вывода пользователю пустых рабочих наборов
        /// </summary>
        private static List<string> GetEmptyWorksets(Document doc)
        {
            List<string> emptyWorksetsNames = new List<string>();
            if (!doc.IsWorkshared) return null;

            List<Workset> wids = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .ToList();

            foreach (Workset w in wids)
            {
                ElementWorksetFilter wfilter = new ElementWorksetFilter(w.Id);
                FilteredElementCollector col = new FilteredElementCollector(doc).WherePasses(wfilter);
                if (col.GetElementCount() == 0)
                {
                    emptyWorksetsNames.Add(w.Name);
                }
            }

            return emptyWorksetsNames;
        }

        /// <summary>
        /// Метод для создания рабочего набора
        /// </summary>
        private static Workset CreateNewWorkset(Document doc, string name)
        {
            bool isUnique = WorksetTable.IsWorksetNameUnique(doc, name);
            if (isUnique)
                Workset.Create(doc, name);

            Workset linkWorkset = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .Where(w => w.Name == name)
                .First();

            return linkWorkset;
        }
    }
}
