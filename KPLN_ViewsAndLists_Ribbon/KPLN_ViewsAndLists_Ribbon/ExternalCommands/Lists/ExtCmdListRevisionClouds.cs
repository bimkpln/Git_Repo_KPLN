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
            Document doc = commandData.Application.ActiveUIDocument.Document;

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
                    List<Revision> revisionsOnSheet = GetRevisionsOnSheet(sheet);

                    if (revisionsOnSheet.Count > 0)
                    {
                        Parameter statusParam = sheet.LookupParameter(StatusParameterName);
                        Parameter noteParam = sheet.LookupParameter(NoteParameterName);

                        int statusCode = statusParam.AsInteger();
                        string noteValue = GetRevisionNoteValue(revisionsOnSheet, statusCode);

                        noteParam.Set(noteValue);
                    }

                    foreach (string parameterName in CloudCountParameterNames)
                    {
                        sheet.LookupParameter(parameterName).Set("");
                    }

                    string[] cloudsCounts = GetCloudsCountsForStampRows(sheet);

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
                problems.Add("- \"" + parameterName + "\" имеет неправильный тип.");
            }

            if (mustBeWritable && parameter.IsReadOnly)
            {
                problems.Add("- \"" + parameterName + "\" доступен только для чтения.");
            }
        }

        private static bool IsParameterBoundToSheets(Document doc, string parameterName)
        {
            Category sheetsCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets);

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

        private static List<Revision> GetRevisionsOnSheet(ViewSheet sheet)
        {
            Document doc = sheet.Document;

            return sheet.GetAllRevisionIds()
                .Select(id => doc.GetElement(id))
                .OfType<Revision>()
                .OrderBy(revision => revision.SequenceNumber)
                .ToList();
        }

        private static string GetRevisionNoteValue(List<Revision> revisionsOnSheet, int statusCode)
        {
            List<string> result = new List<string>();
            string codeString = Math.Abs(statusCode).ToString().PadLeft(4, '0');

            int skipCount = Math.Max(0, revisionsOnSheet.Count - 4);

            List<Revision> revisionsForStamp = revisionsOnSheet
                .Skip(skipCount)
                .ToList();

            for (int i = 0; i < revisionsForStamp.Count; i++)
            {
                Revision revision = revisionsForStamp[i];
                int revisionNumber = revision.SequenceNumber;

                string statusLabel = GetStatusLabelByRow(codeString, i);

                if (string.IsNullOrWhiteSpace(statusLabel))
                    result.Add("Изм. " + revisionNumber);
                else
                    result.Add("Изм. " + revisionNumber + "(" + statusLabel + ")");
            }

            return string.Join(", ", result);
        }

        private static string GetStatusLabelByRow(string codeString, int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= codeString.Length)
                return null;

            char code = codeString[rowIndex];

            if (code == '2')
                return "Зам.";

            if (code == '3')
                return "Нов.";

            return null;
        }


        private static string[] GetCloudsCountsForStampRows(ViewSheet sheet)
        {
            List<Revision> revisionsOnSheet = GetRevisionsOnSheet(sheet);
            List<RevisionCloud> cloudsOnSheet = GetRevisionCloudsOnSheet(sheet);

            int skipCount = Math.Max(0, revisionsOnSheet.Count - CloudCountParameterNames.Length);

            List<Revision> revisionsForStamp = revisionsOnSheet
                .Skip(skipCount)
                .ToList();

            List<string> result = new List<string>();

            foreach (Revision revision in revisionsForStamp)
            {
                int cloudsCount = cloudsOnSheet
                    .Count(cloud => cloud.RevisionId == revision.Id);

                result.Add(cloudsCount == 0 ? "-" : cloudsCount.ToString());
            }

            return result.ToArray();
        }

        private static List<RevisionCloud> GetRevisionCloudsOnSheet(ViewSheet sheet)
        {
            Document doc = sheet.Document;
            HashSet<ElementId> viewIds = GetViewIdsOnSheet(sheet);
            Dictionary<ElementId, RevisionCloud> cloudsById = new Dictionary<ElementId, RevisionCloud>();

            foreach (ElementId viewId in viewIds)
            {
                try
                {
                    List<RevisionCloud> clouds = new FilteredElementCollector(doc, viewId)
                        .OfClass(typeof(RevisionCloud))
                        .Cast<RevisionCloud>()
                        .ToList();

                    foreach (RevisionCloud cloud in clouds)
                    {
                        if (!cloudsById.ContainsKey(cloud.Id))
                            cloudsById.Add(cloud.Id, cloud);
                    }
                }
                catch
                {
                    // Некоторые типы видов Revit не поддерживают FilteredElementCollector по viewId.
                }
            }

            return cloudsById.Values.ToList();
        }

        private static HashSet<ElementId> GetViewIdsOnSheet(ViewSheet sheet)
        {
            Document doc = sheet.Document;
            HashSet<ElementId> viewIds = new HashSet<ElementId>();

            viewIds.Add(sheet.Id);

            foreach (ElementId viewportId in sheet.GetAllViewports())
            {
                Viewport viewport = doc.GetElement(viewportId) as Viewport;

                if (viewport != null && viewport.ViewId != ElementId.InvalidElementId)
                    viewIds.Add(viewport.ViewId);
            }

            return viewIds;
        }
    }
}