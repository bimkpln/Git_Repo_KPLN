using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Lists
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmdListRevisionClouds : IExternalCommand
    {
        private const string StatusParameterName = "Ш.ШифрСтатусЛиста";
        private const string NoteParameterName = "Ш.ПримечаниеЛиста";

        private static readonly string[] CloudCountParameterNames =
        {
            "Ш.КолвоУч1Текст",
            "Ш.КолвоУч2Текст",
            "Ш.КолвоУч3Текст",
            "Ш.КолвоУч4Текст"
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(sheet => !sheet.IsTemplate)
                .ToList();

            if (sheets.Count == 0)
            {
                TaskDialog.Show("KPLN. Внимание", "В проекте нет ни одного листа.");
                return Result.Cancelled;
            }

            List<string> parameterProblems = GetRequiredParameterProblems(doc, sheets[0]);

            if (parameterProblems.Count > 0)
            {
                string text =
                    "Команда не запущена, потому что для категории \"Листы\" не настроены необходимые параметры:"
                    + Environment.NewLine
                    + Environment.NewLine
                    + string.Join(Environment.NewLine, parameterProblems);

                TaskDialog.Show("KPLN. Внимание", text);
                return Result.Cancelled;
            }

            using (Transaction trans = new Transaction(doc, "KPLN. Обновление ИЗМов"))
            {
                trans.Start();

                foreach (ViewSheet sheet in sheets)
                {
                    List<ElementId> revisionIds = sheet.GetAllRevisionIds().ToList();

                    if (revisionIds.Count == 0)
                        continue;

                    Parameter statusParam = sheet.LookupParameter(StatusParameterName);
                    Parameter targetParam = sheet.LookupParameter(NoteParameterName);

                    int code = statusParam.AsInteger();
                    string fieldValue = GetRevisionStatusString(code);

                    if (!string.IsNullOrWhiteSpace(fieldValue))
                    {
                        targetParam.Set(fieldValue);
                    }
                    else
                    {
                        targetParam.Set(GetRevisionListString(revisionIds.Count));
                    }
                }

                foreach (ViewSheet sheet in sheets)
                {
                    string[] cloudsCounts = GetCloudsCountOnSheet(sheet);

                    foreach (string parameterName in CloudCountParameterNames)
                    {
                        sheet.LookupParameter(parameterName).Set("");
                    }

                    for (int i = 0; i < cloudsCounts.Length && i < CloudCountParameterNames.Length; i++)
                    {
                        sheet.LookupParameter(CloudCountParameterNames[i]).Set(cloudsCounts[i]);
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("KPLN. Информация", "Данные об ИЗМах обновлены.");
            return Result.Succeeded;
        }

        private static List<string> GetRequiredParameterProblems(Document doc, ViewSheet sampleSheet)
        {
            List<string> problems = new List<string>();

            CheckRequiredParameter(doc, sampleSheet, StatusParameterName, StorageType.Integer, false, problems);
            CheckRequiredParameter(doc, sampleSheet, NoteParameterName, StorageType.String, true, problems);

            foreach (string parameterName in CloudCountParameterNames)
            {
                CheckRequiredParameter(doc, sampleSheet, parameterName, StorageType.String, true, problems);
            }

            return problems;
        }

        private static void CheckRequiredParameter(
            Document doc,
            ViewSheet sampleSheet,
            string parameterName,
            StorageType expectedStorageType,
            bool mustBeWritable,
            List<string> problems)
        {
            if (!IsParameterBoundToSheets(doc, parameterName))
            {
                problems.Add("- \"" + parameterName + "\" не добавлен к категории \"Листы\".");
                return;
            }

            Parameter parameter = sampleSheet.LookupParameter(parameterName);

            if (parameter == null)
            {
                problems.Add("- \"" + parameterName + "\" добавлен в проект, но не найден на объекте листа.");
                return;
            }

            if (parameter.StorageType != expectedStorageType)
            {
                problems.Add(
                    "- \"" + parameterName + "\" имеет неправильный тип. Ожидается: "
                    + GetStorageTypeName(expectedStorageType)
                    + ".");
            }

            if (mustBeWritable && parameter.IsReadOnly)
            {
                problems.Add("- \"" + parameterName + "\" доступен только для чтения.");
            }
        }

        private static bool IsParameterBoundToSheets(Document doc, string parameterName)
        {
            Category sheetsCategory = Category.GetCategory(doc, BuiltInCategory.OST_Sheets);

            if (sheetsCategory == null)
                return false;

            DefinitionBindingMapIterator iterator = doc.ParameterBindings.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                Definition definition = iterator.Key;

                if (definition == null || definition.Name != parameterName)
                    continue;

                ElementBinding binding = iterator.Current as ElementBinding;

                if (binding == null)
                    continue;

                CategorySetIterator categoryIterator = binding.Categories.ForwardIterator();

                while (categoryIterator.MoveNext())
                {
                    Category category = categoryIterator.Current as Category;

                    if (category != null && category.Id.Equals(sheetsCategory.Id))
                        return true;
                }
            }

            return false;
        }

        private static string GetStorageTypeName(StorageType storageType)
        {
            if (storageType == StorageType.String)
                return "текстовый параметр";

            if (storageType == StorageType.Integer)
                return "целочисленный параметр";

            if (storageType == StorageType.Double)
                return "числовой параметр";

            if (storageType == StorageType.ElementId)
                return "ElementId";

            return storageType.ToString();
        }

        private static string GetRevisionListString(int revisionCount)
        {
            List<string> result = new List<string>();

            for (int i = 0; i < revisionCount; i++)
            {
                result.Add("Изм. " + (i + 1));
            }

            return string.Join(", ", result);
        }

        private string GetRevisionStatusString(int code)
        {
            List<string> result = new List<string>();
            string codeStr = code.ToString().PadLeft(4, '0');

            for (int i = 0; i < codeStr.Length && i < 4; i++)
            {
                char c = codeStr[i];
                string label = null;

                if (c == '2')
                    label = "Зам.";
                else if (c == '3')
                    label = "Нов.";

                if (label != null)
                {
                    int revNum = i + 1;
                    result.Add("Изм. " + revNum + "(" + label + ")");
                }
            }

            return string.Join(", ", result);
        }

        private string[] GetCloudsCountOnSheet(ViewSheet sheet)
        {
            Document doc = sheet.Document;

            List<RevisionCloud> clouds = new FilteredElementCollector(doc)
                .OfClass(typeof(RevisionCloud))
                .Cast<RevisionCloud>()
                .Where(cloud => cloud.OwnerViewId == sheet.Id)
                .ToList();

            List<Revision> allRevisionsOnSheet = sheet.GetAllRevisionIds()
                .Select(id => doc.GetElement(id))
                .OfType<Revision>()
                .OrderBy(revision => revision.IssuedTo)
                .ToList();

            int revisionsCount = allRevisionsOnSheet.Count;
            List<Revision> lastRevisionsOnSheet;

            if (revisionsCount > 4)
            {
                lastRevisionsOnSheet = allRevisionsOnSheet.GetRange(revisionsCount - 4, 4);
            }
            else
            {
                lastRevisionsOnSheet = allRevisionsOnSheet;
            }

            string[] cloudsCounts = new string[lastRevisionsOnSheet.Count];

            for (int i = 0; i < lastRevisionsOnSheet.Count; i++)
            {
                Revision revision = lastRevisionsOnSheet[i];

                int curRevCloudsCount = clouds
                    .Where(cloud => cloud.RevisionId == revision.Id)
                    .Count();

                cloudsCounts[i] = curRevCloudsCount == 0
                    ? "-"
                    : curRevCloudsCount.ToString();
            }

            return cloudsCounts;
        }
    }
}