using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ApartmentManager.Common;
using KPLN_ApartmentManager.Forms;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace KPLN_ApartmentManager.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler : IExternalEventHandler
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";
        private const string ApartmentInstanceMarker = "[KPLN_APT_INSTANCE]";

        private static readonly Dictionary<string, string> _familyNameByPathCache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private ApartmentManagerWindow _window;

        private static void ApplyApartmentFailureHandling(Transaction transaction)
        {
            ApplyApartmentFailureHandling(transaction, new ApartmentWarningSuppressor());
        }

        private static void ApplyApartmentFailureHandling(Transaction transaction, IFailuresPreprocessor preprocessor)
        {
            if (transaction == null)
                return;

            FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
            if (options == null)
                return;

            options.SetFailuresPreprocessor(preprocessor);
            transaction.SetFailureHandlingOptions(options);
        }

        private class ApartmentWarningSuppressor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                if (failuresAccessor == null)
                    return FailureProcessingResult.Continue;

                IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
                if (failures == null || failures.Count == 0)
                    return FailureProcessingResult.Continue;

                foreach (FailureMessageAccessor failure in failures.ToList())
                {
                    if (failure == null || failure.GetSeverity() != FailureSeverity.Warning)
                        continue;

                    string description = null;
                    try
                    {
                        description = failure.GetDescriptionText();
                    }
                    catch
                    {
                        description = null;
                    }

                    if (ShouldSuppressApartmentWarning(description))
                        failuresAccessor.DeleteWarning(failure);
                }

                return FailureProcessingResult.Continue;
            }
        }

        private static bool ShouldSuppressApartmentWarning(string description)
        {
            if (string.IsNullOrWhiteSpace(description))
                return false;

            string normalized = description
                .Replace('ё', 'е')
                .ToLowerInvariant();

            return normalized.Contains("слегка отклони") ||
                   normalized.Contains("slightly off axis") ||
                   (normalized.Contains("стена") &&
                    normalized.Contains("линия-разделитель помещений") &&
                    normalized.Contains("перекры")) ||
                   (normalized.Contains("wall") &&
                    normalized.Contains("room separation") &&
                    normalized.Contains("overlap"));
        }

        public string GetName()
        {
            return "KPLN. Менеджер квартир";
        }

        public void AttachWindow(ApartmentManagerWindow window)
        {
            _window = window;
        }

        public void DetachWindow()
        {
            _window = null;
        }

        private void RestoreWindow()
        {
            if (_window == null)
                return;

            if (_window.Dispatcher == null)
                return;

            _window.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_window == null)
                    return;

                if (!_window.IsVisible)
                    return;

                if (_window.WindowState == WindowState.Minimized)
                    _window.WindowState = WindowState.Normal;

                _window.Show();
                _window.Activate();

                _window.Topmost = true;
                _window.Topmost = false;
                _window.Focus();
            }), DispatcherPriority.Background);
        }

        private class ApartmentFamilyLoadOptions : IFamilyLoadOptions
        {
            private readonly ApartmentFamilyLoadDiagnostic _diagnostic;

            public ApartmentFamilyLoadOptions(ApartmentFamilyLoadDiagnostic diagnostic = null)
            {
                _diagnostic = diagnostic;
            }

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = false;
                if (_diagnostic != null)
                    _diagnostic.Add(
                        "Revit запросил действие для уже загруженного семейства: " +
                        "familyInUse = " + familyInUse +
                        ", overwriteParameterValues = false, продолжить = true.");
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = false;
                if (_diagnostic != null)
                {
                    _diagnostic.AddSharedFamilyDecision(sharedFamily, familyInUse, source, overwriteParameterValues);
                    _diagnostic.Add(
                        "Revit запросил действие для вложенного shared-семейства '" +
                        GetFamilyNameForDiagnostic(sharedFamily) +
                        "': familyInUse = " + familyInUse +
                        ", source = Family, overwriteParameterValues = false, продолжить = true.");
                }
                return true;
            }
        }

        private class ApartmentFamilyReloadOptions : IFamilyLoadOptions
        {
            private readonly ApartmentFamilyLoadDiagnostic _diagnostic;

            public ApartmentFamilyReloadOptions(ApartmentFamilyLoadDiagnostic diagnostic = null)
            {
                _diagnostic = diagnostic;
            }

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                if (_diagnostic != null)
                    _diagnostic.Add(
                        "Revit запросил действие для уже загруженного семейства: " +
                        "familyInUse = " + familyInUse +
                        ", overwriteParameterValues = true, продолжить = true.");
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                if (_diagnostic != null)
                {
                    _diagnostic.AddSharedFamilyDecision(sharedFamily, familyInUse, source, overwriteParameterValues);
                    _diagnostic.Add(
                        "Revit запросил действие для вложенного shared-семейства '" +
                        GetFamilyNameForDiagnostic(sharedFamily) +
                        "': familyInUse = " + familyInUse +
                        ", source = Family, overwriteParameterValues = true, продолжить = true.");
                }
                return true;
            }
        }

        private class ApartmentFailureDiagnosticCollector : IFailuresPreprocessor
        {
            private readonly List<string> _messages = new List<string>();

            public IList<string> Messages
            {
                get { return _messages; }
            }

            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                if (failuresAccessor == null)
                    return FailureProcessingResult.Continue;

                IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
                if (failures == null || failures.Count == 0)
                    return FailureProcessingResult.Continue;

                foreach (FailureMessageAccessor failure in failures.ToList())
                {
                    if (failure == null)
                        continue;

                    string description = GetFailureDescription(failure);
                    string message = GetFailureSeverityName(failure) + ": " + FormatDiagnosticValue(description);

                    ICollection<ElementId> elementIds = GetFailureElementIds(failure);
                    if (elementIds != null && elementIds.Count > 0)
                    {
                        message += " Элементы: " + string.Join(
                            ", ",
                            elementIds
                                .Where(x => x != null && x != ElementId.InvalidElementId)
                                .Take(8)
                                .Select(x => IDHelper.ElIdValue(x).ToString())
                                .ToArray());

                        if (elementIds.Count > 8)
                            message += ", ...";
                    }

                    if (!_messages.Any(x => string.Equals(x, message, StringComparison.OrdinalIgnoreCase)))
                        _messages.Add(message);

                    if (IsFailureWarning(failure) && ShouldSuppressApartmentWarning(description))
                    {
                        failuresAccessor.DeleteWarning(failure);
                    }
                }

                return FailureProcessingResult.Continue;
            }

            private static string GetFailureDescription(FailureMessageAccessor failure)
            {
                try
                {
                    return failure.GetDescriptionText();
                }
                catch
                {
                    return null;
                }
            }

            private static string GetFailureSeverityName(FailureMessageAccessor failure)
            {
                try
                {
                    FailureSeverity severity = failure.GetSeverity();
                    if (severity == FailureSeverity.Warning)
                        return "Предупреждение Revit";
                    if (severity == FailureSeverity.Error)
                        return "Ошибка Revit";

                    return "Сообщение Revit";
                }
                catch
                {
                    return "Сообщение Revit";
                }
            }

            private static ICollection<ElementId> GetFailureElementIds(FailureMessageAccessor failure)
            {
                try
                {
                    return failure.GetFailingElementIds();
                }
                catch
                {
                    return null;
                }
            }

            private static bool IsFailureWarning(FailureMessageAccessor failure)
            {
                try
                {
                    return failure.GetSeverity() == FailureSeverity.Warning;
                }
                catch
                {
                    return false;
                }
            }
        }

        private class ApartmentFamilyLoadDiagnostic
        {
            private readonly List<string> _messages = new List<string>();
            private readonly List<string> _exceptions = new List<string>();
            private readonly List<string> _failureMessages = new List<string>();
            private readonly List<string> _fileProblemMessages = new List<string>();
            private readonly List<string> _hostConflictMessages = new List<string>();
            private readonly List<string> _sharedFamilyDecisions = new List<string>();
            private bool? _loadFamilyResult;
            private string _loadedFamilyName;

            public string FamilyPath { get; private set; }

            public bool HasFileProblem
            {
                get { return _fileProblemMessages.Count > 0; }
            }

            public bool HasBlockingProblem
            {
                get
                {
                    return _fileProblemMessages.Count > 0 ||
                           _hostConflictMessages.Count > 0 ||
                           _failureMessages.Count > 0 ||
                           _exceptions.Count > 0;
                }
            }

            public ApartmentFamilyLoadDiagnostic(string familyPath)
            {
                FamilyPath = familyPath;
            }

            public void Add(string message)
            {
                if (string.IsNullOrWhiteSpace(message))
                    return;

                _messages.Add(message.Trim());
            }

            public void AddException(string stage, Exception ex)
            {
                if (ex == null)
                    return;

                string message = stage + ": " + ex.GetType().Name + ": " + ex.Message;
                _exceptions.Add(message);
                Add(message);
            }

            public void AddFamilyFileOpenException(Exception ex)
            {
                AddFamilyFileOpenException("Ошибка открытия файла семейства через OpenDocumentFile", ex);
            }

            public void AddFamilyFileOpenException(string stage, Exception ex)
            {
                if (ex == null)
                    return;

                if (IsSavedByLaterRevitVersionException(ex))
                {
                    AddFileProblem(
                        "Файл семейства сохранён в более новой версии Revit. Текущий Revit не может его открыть или загрузить. Сообщение Revit: " +
                        ex.Message);
                    return;
                }

                if (string.Equals(ex.GetType().Name, "CorruptModelException", StringComparison.OrdinalIgnoreCase))
                {
                    AddFileProblem("Revit не смог открыть файл семейства. Сообщение Revit: " + ex.Message);
                    return;
                }

                AddException(stage, ex);
            }

            public void AddFileProblem(string message)
            {
                if (string.IsNullOrWhiteSpace(message))
                    return;

                string trimmed = message.Trim();
                if (!_fileProblemMessages.Any(x => string.Equals(x, trimmed, StringComparison.OrdinalIgnoreCase)))
                    _fileProblemMessages.Add(trimmed);
            }

            public void AddFailureMessages(IEnumerable<string> messages)
            {
                if (messages == null)
                    return;

                foreach (string message in messages)
                {
                    if (string.IsNullOrWhiteSpace(message))
                        continue;

                    string trimmed = message.Trim();
                    if (!_failureMessages.Any(x => string.Equals(x, trimmed, StringComparison.OrdinalIgnoreCase)))
                        _failureMessages.Add(trimmed);

                    Add("Failure API: " + trimmed);
                }
            }

            public void AddLoadFamilyResult(bool loaded, Family loadedFamily)
            {
                _loadFamilyResult = loaded;
                _loadedFamilyName = loadedFamily == null ? null : GetFamilyNameForDiagnostic(loadedFamily);
            }

            public void AddHostConflict(string message)
            {
                if (string.IsNullOrWhiteSpace(message))
                    return;

                string trimmed = message.Trim();
                if (!_hostConflictMessages.Any(x => string.Equals(x, trimmed, StringComparison.OrdinalIgnoreCase)))
                    _hostConflictMessages.Add(trimmed);
            }

            public void AddSharedFamilyDecision(Family sharedFamily, bool familyInUse, FamilySource source, bool overwriteParameterValues)
            {
                string familyName = GetFamilyNameForDiagnostic(sharedFamily);
                string placementType = GetFamilyPlacementTypeForDiagnostic(sharedFamily);
                string sourceName = source == FamilySource.Family ? "из файла" : "из проекта";

                string message =
                    familyName +
                    " (" + placementType +
                    ", используется в проекте = " + (familyInUse ? "да" : "нет") +
                    ", выбрано = " + sourceName +
                    ", параметры = " + (overwriteParameterValues ? "перезаписать" : "не перезаписывать") +
                    ")";

                if (!_sharedFamilyDecisions.Any(x => string.Equals(x, message, StringComparison.OrdinalIgnoreCase)))
                    _sharedFamilyDecisions.Add(message);
            }

            public string BuildReport(string title)
            {
                List<string> lines = new List<string>();

                if (!string.IsNullOrWhiteSpace(title))
                    lines.Add(title.Trim());

                if (!string.IsNullOrWhiteSpace(FamilyPath))
                {
                    lines.Add("Файл:");
                    lines.Add(FamilyPath);
                }

                bool hasPrimaryCause = _fileProblemMessages.Count > 0 || _hostConflictMessages.Count > 0;

                if (_fileProblemMessages.Count > 0)
                {
                    lines.Add("");
                    lines.Add("Проблема файла:");

                    foreach (string message in _fileProblemMessages.Take(3))
                        lines.Add("- " + message);

                    if (_fileProblemMessages.Count > 3)
                        lines.Add("- ...");
                }

                if (_hostConflictMessages.Count > 0)
                {
                    lines.Add("");
                    lines.Add("Найден конфликт основы:");

                    foreach (string message in _hostConflictMessages.Take(4))
                        lines.Add("- " + message);

                    if (_hostConflictMessages.Count > 4)
                        lines.Add("- ...");
                }

                if (_loadFamilyResult.HasValue)
                {
                    lines.Add("");
                    if (_loadFamilyResult.Value)
                    {
                        lines.Add("Результат: семейство загружено.");
                        if (!string.IsNullOrWhiteSpace(_loadedFamilyName))
                            lines.Add("Загружено: " + _loadedFamilyName);
                    }
                    else
                    {
                        lines.Add("Технически:");
                        lines.Add("doc.LoadFamily вернул false.");
                    }
                }

                if (!hasPrimaryCause && _sharedFamilyDecisions.Count > 0)
                {
                    lines.Add("");
                    lines.Add("Что удалось получить из Revit API:");
                    lines.Add("Revit запросил решение по вложенным shared-семействам, но точный текст ручного окна не отдал.");

                    foreach (string message in _sharedFamilyDecisions.Take(4))
                        lines.Add("- " + message);

                    if (_sharedFamilyDecisions.Count > 4)
                        lines.Add("- ...");
                }

                if (!hasPrimaryCause && _failureMessages.Count > 0)
                {
                    lines.Add("");
                    lines.Add("Сообщения Revit:");

                    foreach (string message in _failureMessages.Take(5))
                        lines.Add("- " + message);

                    if (_failureMessages.Count > 5)
                        lines.Add("- ...");
                }

                if (!hasPrimaryCause && _exceptions.Count > 0)
                {
                    lines.Add("");
                    lines.Add("Исключения:");

                    foreach (string message in _exceptions.Take(3))
                        lines.Add("- " + message);

                    if (_exceptions.Count > 3)
                        lines.Add("- ...");
                }

                if (!_loadFamilyResult.HasValue && !hasPrimaryCause && _sharedFamilyDecisions.Count == 0 && _failureMessages.Count == 0 && _exceptions.Count == 0 && _messages.Count > 0)
                {
                    lines.Add("");
                    lines.Add("Технические детали:");

                    foreach (string message in _messages.Take(4))
                        lines.Add("- " + message);

                    if (_messages.Count > 4)
                        lines.Add("- ...");
                }

                return LimitDialogText(string.Join("\n", lines.ToArray()), 3500);
            }
        }

        private class FamilyHostDiagnosticInfo
        {
            public string Name { get; set; }
            public string PlacementTypeName { get; set; }
            public bool? HasHost { get; set; }
        }

        private static bool IsSavedByLaterRevitVersionException(Exception ex)
        {
            if (ex == null)
                return false;

            string typeName = ex.GetType().Name ?? "";
            string message = ex.Message ?? "";

            return
                typeName.IndexOf("CorruptModelException", StringComparison.OrdinalIgnoreCase) >= 0 &&
                (message.IndexOf("saved by a later version", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 message.IndexOf("later version", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 message.IndexOf("более новой версии", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 message.IndexOf("более поздней версии", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static string GetFamilyNameForDiagnostic(Family family)
        {
            if (family == null)
                return "<неизвестно>";

            try
            {
                return !string.IsNullOrWhiteSpace(family.Name)
                    ? family.Name.Trim()
                    : "<без имени>";
            }
            catch
            {
                return "<не удалось прочитать имя>";
            }
        }

        private static string GetFamilyPlacementTypeForDiagnostic(Family family)
        {
            if (family == null)
                return "тип размещения неизвестен";

            string placementName = GetFamilyPlacementTypeNameForDiagnostic(family);
            if (string.IsNullOrWhiteSpace(placementName))
                return "тип размещения не прочитан";

            bool? hasHost = IsHostedPlacementTypeName(placementName);
            if (hasHost == true)
                return "с основой: " + placementName;

            if (hasHost == false)
                return "без основы: " + placementName;

            return "тип размещения: " + placementName;
        }

        private static string GetFamilyPlacementTypeNameForDiagnostic(Family family)
        {
            if (family == null)
                return null;

            try
            {
                FamilyPlacementType placementType = family.FamilyPlacementType;
                return placementType.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static bool? IsHostedPlacementTypeName(string placementTypeName)
        {
            if (string.IsNullOrWhiteSpace(placementTypeName))
                return null;

            if (placementTypeName.IndexOf("Hosted", StringComparison.OrdinalIgnoreCase) >= 0 ||
                placementTypeName.IndexOf("FaceBased", StringComparison.OrdinalIgnoreCase) >= 0 ||
                placementTypeName.IndexOf("WallBased", StringComparison.OrdinalIgnoreCase) >= 0 ||
                placementTypeName.IndexOf("CeilingBased", StringComparison.OrdinalIgnoreCase) >= 0 ||
                placementTypeName.IndexOf("FloorBased", StringComparison.OrdinalIgnoreCase) >= 0 ||
                placementTypeName.IndexOf("RoofBased", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        private static FamilyHostDiagnosticInfo BuildFamilyHostDiagnosticInfo(Family family)
        {
            if (family == null)
                return null;

            string name = GetFamilyNameForDiagnostic(family);
            if (string.IsNullOrWhiteSpace(name) || name.StartsWith("<"))
                return null;

            string placementTypeName = GetFamilyPlacementTypeNameForDiagnostic(family);

            return new FamilyHostDiagnosticInfo
            {
                Name = name,
                PlacementTypeName = string.IsNullOrWhiteSpace(placementTypeName) ? "<не прочитан>" : placementTypeName,
                HasHost = IsHostedPlacementTypeName(placementTypeName)
            };
        }

        private static void AddFamilyHostConflictDiagnostics(
            Document projectDoc,
            string familyPath,
            IEnumerable<Family> projectFamilies,
            ApartmentFamilyLoadDiagnostic diagnostic)
        {
            if (diagnostic == null || projectDoc == null || string.IsNullOrWhiteSpace(familyPath) || !File.Exists(familyPath))
                return;

            if (diagnostic.HasFileProblem)
                return;

            List<Family> projectFamilyList = projectFamilies == null
                ? new FilteredElementCollector(projectDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .ToList()
                : projectFamilies.Where(x => x != null).ToList();

            Dictionary<string, FamilyHostDiagnosticInfo> projectFamilyByName = projectFamilyList
                .Select(BuildFamilyHostDiagnosticInfo)
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Name))
                .GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

            Document familyDoc = null;

            try
            {
                familyDoc = projectDoc.Application.OpenDocumentFile(familyPath);

                IEnumerable<FamilyHostDiagnosticInfo> fileFamilyDefinitions = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Select(BuildFamilyHostDiagnosticInfo);

                IEnumerable<FamilyHostDiagnosticInfo> fileFamiliesFromInstances = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Select(GetFamilyFromInstanceForDiagnostic)
                    .Select(BuildFamilyHostDiagnosticInfo);

                List<FamilyHostDiagnosticInfo> fileFamilies = fileFamilyDefinitions
                    .Concat(fileFamiliesFromInstances)
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Name))
                    .GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                    .Select(x => x.First())
                    .ToList();

                foreach (FamilyHostDiagnosticInfo fileFamily in fileFamilies)
                {
                    FamilyHostDiagnosticInfo projectFamily;
                    if (!projectFamilyByName.TryGetValue(fileFamily.Name, out projectFamily))
                        continue;

                    if (!projectFamily.HasHost.HasValue || !fileFamily.HasHost.HasValue)
                        continue;

                    if (projectFamily.HasHost.Value == fileFamily.HasHost.Value)
                        continue;

                    diagnostic.AddHostConflict(BuildFamilyHostConflictMessage(projectFamily, fileFamily));
                }
            }
            catch (Exception ex)
            {
                diagnostic.AddFamilyFileOpenException(ex);
            }
            finally
            {
                if (familyDoc != null)
                {
                    try
                    {
                        familyDoc.Close(false);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static string BuildFamilyHostConflictMessage(FamilyHostDiagnosticInfo projectFamily, FamilyHostDiagnosticInfo fileFamily)
        {
            string projectHost = projectFamily.HasHost == true ? "с основой" : "без основы";
            string fileHost = fileFamily.HasHost == true ? "с основой" : "без основы";

            return "В проекте уже имеется версия семейства '" + projectFamily.Name + "' " + projectHost +
                   ". Ее невозможно заменить семейством " + fileHost +
                   ". (проект: " + projectFamily.PlacementTypeName +
                   "; файл: " + fileFamily.PlacementTypeName + ")";
        }

        private static Family GetFamilyFromInstanceForDiagnostic(FamilyInstance instance)
        {
            if (instance == null)
                return null;

            try
            {
                return instance.Symbol != null ? instance.Symbol.Family : null;
            }
            catch
            {
                return null;
            }
        }

        private static string BuildFamilySymbolDiagnosticName(FamilySymbol symbol)
        {
            if (symbol == null)
                return "<неизвестно>";

            string familyName = "";
            string typeName = "";

            try
            {
                if (symbol.Family != null && !string.IsNullOrWhiteSpace(symbol.Family.Name))
                    familyName = symbol.Family.Name.Trim();
            }
            catch
            {
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(symbol.Name))
                    typeName = symbol.Name.Trim();
            }
            catch
            {
            }

            if (!string.IsNullOrWhiteSpace(familyName) && !string.IsNullOrWhiteSpace(typeName))
                return familyName + " - " + typeName;

            if (!string.IsNullOrWhiteSpace(typeName))
                return typeName;

            if (!string.IsNullOrWhiteSpace(familyName))
                return familyName;

            return "<без имени>";
        }

        private static string BuildApartmentPlacementFailureReport(
            string familyPath,
            FamilySymbol symbol,
            XYZ insertPoint,
            IEnumerable<string> failureMessages,
            Exception exception)
        {
            List<string> lines = new List<string>();
            lines.Add("Не удалось разместить 2D-семейство квартиры.");

            if (!string.IsNullOrWhiteSpace(familyPath))
            {
                lines.Add("");
                lines.Add("Файл:");
                lines.Add(familyPath);
            }

            lines.Add("");
            lines.Add("Тип:");
            lines.Add(BuildFamilySymbolDiagnosticName(symbol));

            lines.Add("Точка:");
            lines.Add(FormatPointForFamilyLoadDiagnostic(insertPoint));

            if (exception != null)
            {
                lines.Add("");
                lines.Add("Исключение:");
                lines.Add(exception.GetType().Name + ": " + exception.Message);
            }

            if (failureMessages != null)
            {
                List<string> messages = failureMessages
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (messages.Count > 0)
                {
                    lines.Add("");
                    lines.Add("Сообщения Revit:");

                    foreach (string message in messages)
                        lines.Add("- " + message);
                }
            }

            return LimitDialogText(string.Join("\n", lines.ToArray()), 6000);
        }

        private static string FormatPointForFamilyLoadDiagnostic(XYZ point)
        {
            if (point == null)
                return "<нет>";

            try
            {
                return "(" +
                       Math.Round(IDHelper.ConvertInternalToMm(point.X)).ToString() + "; " +
                       Math.Round(IDHelper.ConvertInternalToMm(point.Y)).ToString() + "; " +
                       Math.Round(IDHelper.ConvertInternalToMm(point.Z)).ToString() +
                       ") мм";
            }
            catch
            {
                return "<не удалось отформатировать>";
            }
        }

        private static string LimitDialogText(string text, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(text) || maxLength <= 0 || text.Length <= maxLength)
                return text;

            return text.Substring(0, maxLength) + "\n...";
        }

        private class ApartmentFamilyReloadCandidate
        {
            public string FamilyPath { get; set; }
            public string FamilyName { get; set; }
        }

        private class ExistingWallLineInfo
        {
            public ElementId WallId { get; set; }
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
            public double ThicknessInternal { get; set; }
        }

        private class PreparedApartmentWalls
        {
            public ElementId ApartmentId { get; set; }
            public WallType WallType { get; set; }
            public int ThicknessMm { get; set; }
            public List<Line> AxisLines { get; set; }
            public WallType ShaftWallType { get; set; }
            public List<Line> ShaftAxisLines { get; set; }
            public WallType LoggiaWallType { get; set; }
            public List<Line> LoggiaAxisLines { get; set; }
            public List<Line> RoomSeparatorLines { get; set; }
        }

        private class ApartmentProcessState
        {
            public ElementId ApartmentId { get; set; }
            public string ApartmentFamilyName { get; set; }
            public ElementId NavigationElementId { get; set; }
            public List<ElementId> NavigationElementIds { get; set; }
            public List<ElementId> CreatedElementIds { get; set; }
            public Apartment2DRestoreInfo Restore2DInfo { get; set; }
            public bool HasPreparedWalls { get; set; }
            public bool HasCreatedWalls { get; set; }
            public bool HasCreatedRooms { get; set; }
            public bool HasInstalledDoors { get; set; }
            public bool HasInstalledWindows { get; set; }
            public int FoundEntranceDoorsCount { get; set; }
            public int InstalledEntranceDoorsCount { get; set; }
            public int FoundRoomSeparatorsCount { get; set; }
            public int CreatedRoomSeparatorsCount { get; set; }
            public int SkippedRoomsCount { get; set; }
            public int SkippedWallsCount { get; set; }
            public int SkippedDoorsCount { get; set; }
            public int SkippedWindowsCount { get; set; }
            public List<string> FurnitureErrors { get; set; }
            public List<string> ErrorMessages { get; set; }
            public bool HasRoomAreaMismatch { get; set; }

            public ApartmentProcessState()
            {
                NavigationElementId = ElementId.InvalidElementId;
                NavigationElementIds = new List<ElementId>();
                CreatedElementIds = new List<ElementId>();
                FurnitureErrors = new List<string>();
                ErrorMessages = new List<string>();
            }
        }

        private class OffsetLine2D
        {
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public XYZ Dir { get; set; }
            public double Z { get; set; }
        }

        private class Interval1D
        {
            public double From { get; set; }
            public double To { get; set; }
        }

        private class GenericAxisLineData
        {
            public XYZ Dir { get; set; }
            public XYZ Normal { get; set; }
            public double Offset { get; set; }
            public double Z { get; set; }
            public double From { get; set; }
            public double To { get; set; }
        }

        private class GenericAxisGroupKey : IEquatable<GenericAxisGroupKey>
        {
            public XYZ Dir { get; private set; }
            public double Offset { get; private set; }
            public double Z { get; private set; }

            public GenericAxisGroupKey(XYZ dir, double offset, double z, double tol)
            {
                Dir = new XYZ(RoundTol(dir.X, tol), RoundTol(dir.Y, tol), 0);
                Offset = RoundTol(offset, tol);
                Z = RoundTol(z, tol);
            }

            public bool Equals(GenericAxisGroupKey other)
            {
                if (other == null)
                    return false;

                return
                    Math.Abs(Dir.X - other.Dir.X) < 1e-9 &&
                    Math.Abs(Dir.Y - other.Dir.Y) < 1e-9 &&
                    Math.Abs(Offset - other.Offset) < 1e-9 &&
                    Math.Abs(Z - other.Z) < 1e-9;
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as GenericAxisGroupKey);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + Dir.X.GetHashCode();
                    hash = hash * 23 + Dir.Y.GetHashCode();
                    hash = hash * 23 + Offset.GetHashCode();
                    hash = hash * 23 + Z.GetHashCode();
                    return hash;
                }
            }
        }

        private class FamilyWindowMarker
        {
            public XYZ LocalP0 { get; set; }
            public XYZ LocalP1 { get; set; }
        }

        private class FamilyShaftWallMarker
        {
            public XYZ ProjectP0 { get; set; }
            public XYZ ProjectP1 { get; set; }
        }

        private class FamilyLoggiaWallMarker
        {
            public XYZ ProjectP0 { get; set; }
            public XYZ ProjectP1 { get; set; }
        }

        private class FamilyRoomSeparatorMarker
        {
            public XYZ ProjectP0 { get; set; }
            public XYZ ProjectP1 { get; set; }
        }

        private class HelperLineCandidate
        {
            public XYZ P0 { get; set; }
            public XYZ P1 { get; set; }
            public bool StyleMatched { get; set; }
            public string MatchedName { get; set; }
        }

        private class DoorTypeMirrorEnsureResult
        {
            public bool HasMessage { get; set; }
            public string Message { get; set; }
        }

        private enum DoorOpeningMarker
        {
            None,
            Left,
            Right,
            RightAlt
        }

        private class PreparedDoorPlacement
        {
            public ElementId ApartmentId { get; set; }
            public ElementId Door2DId { get; set; }
            public string RoomCategory { get; set; }
            public int DoorWidthMm { get; set; }
            public string SelectedDoorTypeName { get; set; }
            public FamilySymbol DoorSymbol { get; set; }
            public XYZ InsertPoint { get; set; }
            public FamilyInstance RelatedRoom2D { get; set; }
            public XYZ InteriorReferencePoint { get; set; }
            public XYZ SourceHandDirection { get; set; }
            public XYZ SourceFacingDirection { get; set; }
            public XYZ SourceRoomCalculationSideDirection { get; set; }
            public bool IsEntranceDoor { get; set; }
            public bool UseEntranceDoorPlacement { get; set; }
            public List<string> Diagnostics { get; set; }
        }

        private class PreparedApartmentDoors
        {
            public ElementId ApartmentId { get; set; }
            public List<PreparedDoorPlacement> Doors { get; set; }

            public PreparedApartmentDoors()
            {
                Doors = new List<PreparedDoorPlacement>();
            }
        }

        private class PreparedWindowPlacement
        {
            public ElementId ApartmentId { get; set; }
            public FamilySymbol WindowSymbol { get; set; }
            public Line SourceLine { get; set; }
            public XYZ InsertPoint { get; set; }
            public XYZ ReferenceDirection { get; set; }
            public double SillHeightInternal { get; set; }
            public List<string> Diagnostics { get; set; }
        }

        private class PreparedApartmentWindows
        {
            public ElementId ApartmentId { get; set; }
            public List<PreparedWindowPlacement> Windows { get; set; }

            public PreparedApartmentWindows()
            {
                Windows = new List<PreparedWindowPlacement>();
            }
        }

        private class PreparedRoomPlacement
        {
            public ElementId ApartmentId { get; set; }
            public ElementId SourceRoom2DId { get; set; }
            public string RoomName { get; set; }
            public XYZ InsertPoint { get; set; }
            public List<XYZ> BoundaryVertices { get; set; }
            public double ExpectedAreaInternal { get; set; }
            public double AreaMismatchToleranceSquareMeters { get; set; }
            public bool HasShaftInside { get; set; }
        }

        private class PreparedApartmentRooms
        {
            public ElementId ApartmentId { get; set; }
            public List<PreparedRoomPlacement> Rooms { get; set; }
            public bool HasRoomSeparators { get; set; }
            public bool HasShafts { get; set; }

            public PreparedApartmentRooms()
            {
                Rooms = new List<PreparedRoomPlacement>();
            }
        }

        private class RoomAreaMismatchInfo
        {
            public ElementId ApartmentId { get; set; }
            public string RoomName { get; set; }
            public double ExpectedAreaInternal { get; set; }
            public double ActualAreaInternal { get; set; }
        }

        private enum RequestType
        {
            None,
            PlaceApartment,
            ConvertTo3D,
            RefreshApartmentPresets,
            UpdateApartmentFamilies,
            UpdateApartmentMarks
        }

        private RequestType _requestType = RequestType.None;
        private int _requestedApartmentId;
        private ApartmentPresetData _requestedPresetData;

        public void Execute(UIApplication app)
        {
            try
            {
                UIDocument uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                    throw new Exception("Не найден активный UIDocument.");

                Document doc = uidoc.Document;
                if (doc == null)
                    throw new Exception("Не найден активный документ.");

                switch (_requestType)
                {
                    case RequestType.PlaceApartment:
                        if (ExecutePlaceApartment(uidoc, doc, _requestedApartmentId))
                            MarkApartmentPresetDataStaleInWindow();
                        break;

                    case RequestType.ConvertTo3D:
                        string validationMessage;
                        bool isPresetDataStale;
                        if (!ValidatePresetBeforeConvertTo3D(doc, _requestedPresetData, out validationMessage, out isPresetDataStale))
                        {
                            if (isPresetDataStale)
                                MarkApartmentPresetDataStaleInWindow();

                            TaskDialog.Show("Предупреждение", validationMessage);
                            return;
                        }

                        ExecuteConvertTo3D(uidoc, doc, _requestedPresetData);
                        break;

                    case RequestType.RefreshApartmentPresets:
                        ExecuteRefreshApartmentPresets(doc, _requestedPresetData);
                        break;

                    case RequestType.UpdateApartmentFamilies:
                        ExecuteUpdateApartmentFamilies(doc);
                        break;

                    case RequestType.UpdateApartmentMarks:
                        ExecuteUpdateApartmentMarks(doc);
                        break;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
            }
            finally
            {
                _requestType = RequestType.None;
                _requestedApartmentId = 0;
                _requestedPresetData = null;

                RestoreWindow();
            }
        }
        public void PreparePlaceApartment(int apartmentId)
        {
            _requestType = RequestType.PlaceApartment;
            _requestedApartmentId = apartmentId;
            _requestedPresetData = null;
        }

        public void PrepareConvertTo3D(ApartmentPresetData presetData)
        {
            _requestType = RequestType.ConvertTo3D;
            _requestedApartmentId = 0;
            _requestedPresetData = presetData != null ? presetData.Clone() : null;
        }

        public void PrepareRefreshApartmentPresets(ApartmentPresetData presetData)
        {
            _requestType = RequestType.RefreshApartmentPresets;
            _requestedApartmentId = 0;
            _requestedPresetData = presetData != null ? presetData.Clone() : null;
        }

        public void PrepareUpdateApartmentMarks()
        {
            _requestType = RequestType.UpdateApartmentMarks;
            _requestedApartmentId = 0;
            _requestedPresetData = null;
        }

        public void PrepareUpdateApartmentFamilies()
        {
            _requestType = RequestType.UpdateApartmentFamilies;
            _requestedApartmentId = 0;
            _requestedPresetData = null;
        }

        private void MarkApartmentPresetDataStaleInWindow()
        {
            if (_window == null || _window.Dispatcher == null)
                return;

            _window.Dispatcher.Invoke(new Action(() =>
            {
                _window.MarkApartmentPresetDataStale();
            }));
        }

        private static SQLiteConnection OpenConnection(string dbPath, bool readOnly)
        {
            string cs = readOnly
                ? ("Data Source=" + dbPath + ";Version=3;Read Only=True;")
                : ("Data Source=" + dbPath + ";Version=3;");

            SQLiteConnection con = new SQLiteConnection(cs);
            con.Open();
            return con;
        }

        private static string GetFamilyPathById(int id)
        {
            if (!File.Exists(DbPath))
                throw new FileNotFoundException("Не найдена база данных", DbPath);

            using (SQLiteConnection con = OpenConnection(DbPath, true))
            using (SQLiteCommand cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT FPATH FROM Main WHERE ID = @id LIMIT 1;";
                cmd.Parameters.AddWithValue("@id", id);

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    return null;

                return result.ToString().Trim();
            }
        }

        private static List<string> GetApartmentFamilyPathsFromDb()
        {
            List<string> result = new List<string>();

            if (!File.Exists(DbPath))
                throw new FileNotFoundException("Не найдена база данных", DbPath);

            using (SQLiteConnection con = OpenConnection(DbPath, true))
            using (SQLiteCommand cmd = con.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT DISTINCT FPATH FROM Main " +
                    "WHERE FPATH IS NOT NULL AND TRIM(FPATH) <> '';";

                using (SQLiteDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        string path = r.IsDBNull(0) ? null : r.GetString(0);
                        if (!string.IsNullOrWhiteSpace(path))
                            result.Add(path.Trim());
                    }
                }
            }

            return result
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string GetFamilyNameFromFile(Document projectDoc, string familyPath, ApartmentFamilyLoadDiagnostic diagnostic = null)
        {
            if (projectDoc == null)
            {
                if (diagnostic != null)
                    diagnostic.Add("Не удалось прочитать имя семейства из файла: projectDoc = null.");
                return null;
            }

            if (string.IsNullOrWhiteSpace(familyPath) || !File.Exists(familyPath))
            {
                if (diagnostic != null)
                    diagnostic.Add("Не удалось прочитать имя семейства из файла: файл не найден.");
                return null;
            }

            string cachedName;
            if (_familyNameByPathCache.TryGetValue(familyPath, out cachedName))
            {
                if (diagnostic != null)
                    diagnostic.Add(
                        "Имя семейства из файла взято из кэша: '" +
                        (string.IsNullOrWhiteSpace(cachedName) ? "<пусто>" : cachedName) + "'.");
                return string.IsNullOrWhiteSpace(cachedName) ? null : cachedName;
            }

            Document familyDoc = null;
            string detectedName = null;

            try
            {
                familyDoc = projectDoc.Application.OpenDocumentFile(familyPath);

                Family family = new FilteredElementCollector(familyDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault();

                if (family != null && !string.IsNullOrWhiteSpace(family.Name))
                    detectedName = family.Name.Trim();

                if (diagnostic != null)
                    diagnostic.Add(
                        "Имя семейства из файла: '" +
                        (string.IsNullOrWhiteSpace(detectedName) ? "<не найдено>" : detectedName) + "'.");
            }
            catch (Exception ex)
            {
                if (diagnostic != null)
                    diagnostic.AddFamilyFileOpenException(ex);
            }
            finally
            {
                if (familyDoc != null)
                {
                    try
                    {
                        familyDoc.Close(false);
                    }
                    catch
                    {
                    }
                }
            }

            _familyNameByPathCache[familyPath] = detectedName ?? "";
            return detectedName;
        }

        private static double RoundTol(double value, double tol)
        {
            return Math.Round(value / tol) * tol;
        }

        private static ElementId GetInstanceLevelId(FamilyInstance fi)
        {
            if (fi == null)
                return ElementId.InvalidElementId;

            Parameter p =
                fi.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ??
                fi.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);

            if (p != null && p.StorageType == StorageType.ElementId)
                return p.AsElementId();

            if (fi.LevelId != ElementId.InvalidElementId)
                return fi.LevelId;

            return ElementId.InvalidElementId;
        }

        private static string GetCommentsValue(Element e)
        {
            if (e == null)
                return null;

            List<string> instanceValues = GetCommentParameterValues(e, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            string entranceInstanceValue = instanceValues.FirstOrDefault(IsEntranceDoorComment);
            if (!string.IsNullOrWhiteSpace(entranceInstanceValue))
                return entranceInstanceValue.Trim();

            string firstInstanceValue = instanceValues.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
            if (!string.IsNullOrWhiteSpace(firstInstanceValue))
                return firstInstanceValue.Trim();

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                List<string> typeValues = GetCommentParameterValues(typeElem, BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
                string entranceTypeValue = typeValues.FirstOrDefault(IsEntranceDoorComment);
                if (!string.IsNullOrWhiteSpace(entranceTypeValue))
                    return entranceTypeValue.Trim();

                string firstTypeValue = typeValues.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                if (!string.IsNullOrWhiteSpace(firstTypeValue))
                    return firstTypeValue.Trim();
            }

            return null;
        }

        private static bool IsEntranceDoor2DMarker(FamilyInstance fi)
        {
            if (fi == null || fi.Symbol == null)
                return false;

            return IsEntranceDoor2DTypeName(fi.Symbol.Name);
        }

        private static bool IsEntranceDoor2DTypeName(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
                return false;

            string normalized = typeName
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return normalized.EndsWith("_Входная", StringComparison.OrdinalIgnoreCase);
        }

        private static List<string> GetCommentParameterValues(Element e, BuiltInParameter builtInParameter)
        {
            List<string> values = new List<string>();
            if (e == null)
                return values;

            AddParameterValue(values, e.get_Parameter(builtInParameter));
            AddNamedParameterValues(values, e, "Комментарии");
            AddNamedParameterValues(values, e, "Комментарий");
            AddNamedParameterValues(values, e, "Comments");
            AddNamedParameterValues(values, e, "Comment");

            return values;
        }

        private static void AddNamedParameterValues(List<string> values, Element e, string parameterName)
        {
            if (values == null || e == null || string.IsNullOrWhiteSpace(parameterName))
                return;

            try
            {
                IList<Parameter> parameters = e.GetParameters(parameterName);
                if (parameters != null)
                {
                    foreach (Parameter parameter in parameters)
                        AddParameterValue(values, parameter);
                }
            }
            catch
            {
            }

            try
            {
                AddParameterValue(values, e.LookupParameter(parameterName));
            }
            catch
            {
            }
        }

        private static void AddParameterValue(List<string> values, Parameter parameter)
        {
            if (values == null || parameter == null)
                return;

            string value = null;

            try
            {
                if (parameter.StorageType == StorageType.String)
                    value = parameter.AsString();
                else
                    value = parameter.AsValueString();
            }
            catch
            {
                value = null;
            }

            if (string.IsNullOrWhiteSpace(value))
                return;

            value = value.Trim();
            if (!values.Contains(value))
                values.Add(value);
        }

        private static List<FamilyInstance> GetPlacedApartmentInstancesInDocument(Document doc)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();

            if (doc == null)
                return result;

            IEnumerable<FamilyInstance> instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>();

            foreach (FamilyInstance fi in instances)
            {
                if (fi == null)
                    continue;

                Parameter pComment = fi.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                if (pComment == null)
                    continue;

                string comment = pComment.AsString();
                if (string.IsNullOrWhiteSpace(comment))
                    continue;

                if (!comment.Contains(ApartmentInstanceMarker))
                    continue;

                result.Add(fi);
            }

            return result;
        }

        private static bool SetApartmentAreaParameterValue(FamilyInstance apartmentFi, string parameterName, double areaInternal, List<string> errors)
        {
            if (apartmentFi == null || string.IsNullOrWhiteSpace(parameterName))
                return false;

            Parameter p = apartmentFi.LookupParameter(parameterName);
            if (p == null)
            {
                if (errors != null)
                    errors.Add("У квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " не найден параметр '" + parameterName + "'.");
                return false;
            }

            if (p.IsReadOnly)
            {
                if (errors != null)
                    errors.Add("Параметр '" + parameterName + "' у квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " доступен только для чтения.");
                return false;
            }

            if (p.StorageType != StorageType.Double)
            {
                if (errors != null)
                    errors.Add("Параметр '" + parameterName + "' у квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " имеет некорректный тип.");
                return false;
            }

            p.Set(areaInternal);
            return true;
        }

        private void ExecuteUpdateApartmentMarks(Document doc)
        {
            List<FamilyInstance> apartments = GetPlacedApartmentInstancesInDocument(doc);

            if (apartments.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "В модели не найдены квартиры, размещённые через менеджер.");
                return;
            }

            int updatedCount = 0;
            int skippedCount = 0;
            List<string> errors = new List<string>();

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Обновление марок квартир"))
            {
                t.Start();

                foreach (FamilyInstance apartmentFi in apartments)
                {
                    if (apartmentFi == null)
                        continue;

                    try
                    {
                        List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);

                        double livingAreaInternal = 0.0;
                        double totalAreaInternal = 0.0;

                        foreach (FamilyInstance roomFi in roomInstances)
                        {
                            if (roomFi == null)
                                continue;

                            double roomAreaInternal;
                            if (!TryGetAreaParamFromElementOrType(roomFi, out roomAreaInternal, "КП_Р_Площадь", "КП_Р_ПЛОЩАДЬ"))
                                continue;

                            totalAreaInternal += roomAreaInternal;

                            string roomCategory = GetRoomCategoryLabel(roomFi);
                            if (string.Equals((roomCategory ?? "").Trim(), "Комната", StringComparison.OrdinalIgnoreCase))
                                livingAreaInternal += roomAreaInternal;
                        }

                        bool livingOk = SetApartmentAreaParameterValue(apartmentFi, "КВ_Площадь_Жилая", livingAreaInternal, errors);
                        bool totalOk = SetApartmentAreaParameterValue(apartmentFi, "КВ_Площадь_Общая", totalAreaInternal, errors);

                        if (livingOk && totalOk)
                            updatedCount++;
                        else
                            skippedCount++;
                    }
                    catch (Exception ex)
                    {
                        skippedCount++;
                        errors.Add("Ошибка у квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + ": " + ex.Message);
                    }
                }

                t.Commit();
            }

            string message =
                "Найдено квартир: " + apartments.Count +
                "\nОбновлено: " + updatedCount +
                "\nПропущено: " + skippedCount;

            if (errors.Count > 0)
            {
                List<string> shortErrors = errors.Take(15).ToList();
                message += "\n\nОшибки:\n- " + string.Join("\n- ", shortErrors);

                if (errors.Count > shortErrors.Count)
                    message += "\n- ...";
            }

            TaskDialog.Show("KPLN. Менеджер квартир", message);
        }

        private void ExecuteUpdateApartmentFamilies(Document doc)
        {
            if (doc == null)
                return;

            List<string> familyPaths;
            try
            {
                familyPaths = GetApartmentFamilyPathsFromDb();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "Не удалось прочитать пути семейств из БД:\n" + ex.Message);
                return;
            }

            if (familyPaths.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер квартир", "В БД не найдено путей к семействам квартир.");
                return;
            }

            HashSet<string> loadedFamilyNames = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Name))
                    .Select(x => x.Name.Trim()),
                StringComparer.OrdinalIgnoreCase);

            List<ApartmentFamilyReloadCandidate> candidates = new List<ApartmentFamilyReloadCandidate>();
            List<string> missingFiles = new List<string>();

            foreach (string familyPath in familyPaths)
            {
                if (string.IsNullOrWhiteSpace(familyPath))
                    continue;

                if (!File.Exists(familyPath))
                {
                    missingFiles.Add(familyPath);
                    continue;
                }

                string fileFamilyName = Path.GetFileNameWithoutExtension(familyPath);
                string familyName = fileFamilyName;
                bool isLoaded = !string.IsNullOrWhiteSpace(fileFamilyName) && loadedFamilyNames.Contains(fileFamilyName);

                if (!isLoaded)
                {
                    string realFamilyName = GetFamilyNameFromFile(doc, familyPath);
                    if (!string.IsNullOrWhiteSpace(realFamilyName) && loadedFamilyNames.Contains(realFamilyName))
                    {
                        familyName = realFamilyName.Trim();
                        isLoaded = true;
                    }
                }

                if (!isLoaded)
                    continue;

                candidates.Add(new ApartmentFamilyReloadCandidate
                {
                    FamilyPath = familyPath,
                    FamilyName = familyName
                });
            }

            candidates = candidates
                .GroupBy(x => x.FamilyPath, StringComparer.OrdinalIgnoreCase)
                .Select(x => x.First())
                .OrderBy(x => x.FamilyName)
                .ToList();

            if (candidates.Count == 0)
            {
                string message = "В проекте не найдено загруженных 2D-семейств квартир из БД.";
                if (missingFiles.Count > 0)
                    message += "\n\nНе найдено файлов в БД: " + missingFiles.Count;

                TaskDialog.Show("KPLN. Менеджер квартир", message);
                return;
            }

            int updatedCount = 0;
            int unchangedCount = 0;
            int failedCount = 0;
            List<string> errors = new List<string>();

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Обновление 2D-семейств квартир"))
            {
                t.Start();

                foreach (ApartmentFamilyReloadCandidate candidate in candidates)
                {
                    if (candidate == null || string.IsNullOrWhiteSpace(candidate.FamilyPath))
                        continue;

                    ApartmentFamilyLoadDiagnostic reloadDiagnostic = new ApartmentFamilyLoadDiagnostic(candidate.FamilyPath);

                    try
                    {
                        reloadDiagnostic.Add("Перезагружаем уже загруженное 2D-семейство квартиры '" + candidate.FamilyName + "'.");
                        AddFamilyHostConflictDiagnostics(doc, candidate.FamilyPath, null, reloadDiagnostic);

                        Family loadedFamily;
                        bool loaded = doc.LoadFamily(candidate.FamilyPath, new ApartmentFamilyReloadOptions(reloadDiagnostic), out loadedFamily);
                        reloadDiagnostic.AddLoadFamilyResult(loaded, loadedFamily);
                        reloadDiagnostic.Add(
                            "doc.LoadFamily завершился: loaded = " + loaded +
                            ", loadedFamily = '" + GetFamilyNameForDiagnostic(loadedFamily) + "'.");

                        if (loaded || loadedFamily != null)
                        {
                            updatedCount++;
                        }
                        else if (!reloadDiagnostic.HasBlockingProblem)
                        {
                            unchangedCount++;
                        }
                        else
                        {
                            failedCount++;
                            errors.Add(
                                candidate.FamilyName + ": " +
                                reloadDiagnostic.BuildReport("Revit не перезагрузил семейство."));
                        }
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        reloadDiagnostic.AddFamilyFileOpenException("Исключение во время перезагрузки семейства", ex);
                        errors.Add(
                            candidate.FamilyName + ": " +
                            reloadDiagnostic.BuildReport("Ошибка перезагрузки семейства."));
                    }
                }

                t.Commit();
            }

            if (updatedCount > 0)
                MarkApartmentPresetDataStaleInWindow();

            string resultMessage =
                "Найдено в проекте: " + candidates.Count +
                "\nОбновлено: " + updatedCount +
                "\nБез изменений: " + unchangedCount +
                "\nНе обновлено: " + failedCount;

            if (missingFiles.Count > 0)
                resultMessage += "\nНе найдено файлов из БД: " + missingFiles.Count;

            if (errors.Count > 0)
            {
                List<string> shortErrors = errors
                    .Take(5)
                    .Select(x => LimitDialogText(x, 1200))
                    .ToList();
                resultMessage += "\n\nОшибки:\n- " + string.Join("\n- ", shortErrors);
                if (errors.Count > shortErrors.Count)
                    resultMessage += "\n- ...";
            }

            TaskDialog.Show("KPLN. Менеджер квартир", resultMessage);
        }

        private bool ExecutePlaceApartment(UIDocument uidoc, Document doc, int id)
        {
            ViewPlan floorPlan = doc.ActiveView as ViewPlan;
            if (floorPlan == null || floorPlan.ViewType != ViewType.FloorPlan)
            {
                TaskDialog.Show("Предупреждение", "Откройте план этажа перед размещением квартиры.");
                return false;
            }

            string familyPath = GetFamilyPathById(id);
            if (string.IsNullOrWhiteSpace(familyPath))
            {
                TaskDialog.Show("Ошибка", "Для ID = " + id + " не найден FPATH в базе.");
                return false;
            }

            if (!File.Exists(familyPath))
            {
                TaskDialog.Show("Ошибка", "Файл семейства не найден:\n" + familyPath);
                return false;
            }

            FamilySymbol symbol = null;
            ApartmentFamilyLoadDiagnostic loadDiagnostic = new ApartmentFamilyLoadDiagnostic(familyPath);
            ApartmentFailureDiagnosticCollector failureCollector = new ApartmentFailureDiagnosticCollector();

            try
            {
                using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Загрузка семейства квартиры"))
                {
                    t.Start();
                    ApplyApartmentFailureHandling(t, failureCollector);

                    Family family = LoadOrFindFamily(doc, familyPath, loadDiagnostic);
                    if (family == null)
                        throw new Exception("Не удалось загрузить или найти семейство в проекте.");

                    loadDiagnostic.Add("Итоговое семейство в проекте: '" + GetFamilyNameForDiagnostic(family) + "'.");

                    symbol = GetFirstFamilySymbol(doc, family);
                    if (symbol == null)
                        throw new Exception("В семействе не найден ни один типоразмер.");

                    loadDiagnostic.Add("Первый типоразмер: '" + BuildFamilySymbolDiagnosticName(symbol) + "'.");

                    if (!symbol.IsActive)
                    {
                        loadDiagnostic.Add("Активируем типоразмер.");
                        symbol.Activate();
                        doc.Regenerate();
                    }

                    t.Commit();
                }

                loadDiagnostic.AddFailureMessages(failureCollector.Messages);
            }
            catch (Exception ex)
            {
                loadDiagnostic.AddFailureMessages(failureCollector.Messages);
                loadDiagnostic.AddException("Ошибка на этапе загрузки/активации семейства квартиры", ex);
                TaskDialog.Show(
                    "KPLN. Менеджер квартир",
                    loadDiagnostic.BuildReport("Не удалось загрузить 2D-семейство квартиры."));
                return false;
            }

            int placedCount = 0;

            using (TransactionGroup tg = new TransactionGroup(doc, "KPLN. Менеджер квартир. Размещение семейства квартиры"))
            {
                tg.Start();

                while (true)
                {
                    XYZ insertPoint;

                    try
                    {
                        insertPoint = uidoc.Selection.PickPoint("Укажите точку вставки квартиры. ESC - завершить.");
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        break;
                    }

                    ApartmentFailureDiagnosticCollector placementFailureCollector = new ApartmentFailureDiagnosticCollector();

                    try
                    {
                        using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Размещение семейства квартиры"))
                        {
                            t.Start();
                            ApplyApartmentFailureHandling(t, placementFailureCollector);

                            FamilyInstance placedInstance = PlaceFamilyInstance(doc, floorPlan, symbol, insertPoint);
                            if (placedInstance == null)
                                throw new Exception("Revit не создал экземпляр семейства.");

                            AppendComment(placedInstance, ApartmentInstanceMarker);

                            t.Commit();
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show(
                            "KPLN. Менеджер квартир",
                            BuildApartmentPlacementFailureReport(familyPath, symbol, insertPoint, placementFailureCollector.Messages, ex));
                        break;
                    }

                    placedCount++;
                }

                if (placedCount > 0)
                    tg.Assimilate();
                else
                    tg.RollBack();
            }

            return placedCount > 0;
        }

        private static Family LoadOrFindFamily(Document doc, string familyPath, ApartmentFamilyLoadDiagnostic diagnostic = null)
        {
            if (doc == null)
                throw new ArgumentNullException("doc");

            if (string.IsNullOrWhiteSpace(familyPath))
                throw new ArgumentException("Не задан путь к семейству.", "familyPath");

            if (!File.Exists(familyPath))
                throw new FileNotFoundException("Файл семейства не найден.", familyPath);

            string fileFamilyName = Path.GetFileNameWithoutExtension(familyPath);

            if (diagnostic != null)
            {
                diagnostic.Add("Имя файла без расширения: '" + fileFamilyName + "'.");
                diagnostic.Add("Проверяем, загружено ли семейство в проект.");
            }

            List<Family> existingFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            Family existingByFileName = existingFamilies.FirstOrDefault(f =>
                string.Equals(f.Name, fileFamilyName, StringComparison.OrdinalIgnoreCase));

            if (diagnostic != null)
                diagnostic.Add(
                    existingByFileName != null
                        ? "Семейство уже найдено в проекте по имени файла: '" + GetFamilyNameForDiagnostic(existingByFileName) + "'."
                        : "Семейство по имени файла в проекте не найдено.");

            if (existingByFileName != null)
                return existingByFileName;

            string realFamilyName = GetFamilyNameFromFile(doc, familyPath, diagnostic);

            if (!string.IsNullOrWhiteSpace(realFamilyName))
            {
                Family existingByRealName = existingFamilies.FirstOrDefault(f =>
                    string.Equals(f.Name, realFamilyName, StringComparison.OrdinalIgnoreCase));

                if (diagnostic != null)
                    diagnostic.Add(
                        existingByRealName != null
                            ? "Семейство уже найдено в проекте по имени из файла: '" + GetFamilyNameForDiagnostic(existingByRealName) + "'."
                            : "Семейство по имени из файла в проекте не найдено.");

                if (existingByRealName != null)
                    return existingByRealName;
            }
            else if (diagnostic != null)
            {
                diagnostic.Add("Реальное имя семейства из файла определить не удалось.");
            }

            Family loadedFamily = null;

            AddFamilyHostConflictDiagnostics(doc, familyPath, existingFamilies, diagnostic);

            try
            {
                if (diagnostic != null)
                    diagnostic.Add("Вызываем doc.LoadFamily для загрузки семейства квартиры.");

                bool loaded = doc.LoadFamily(familyPath, new ApartmentFamilyLoadOptions(diagnostic), out loadedFamily);

                if (diagnostic != null)
                {
                    diagnostic.AddLoadFamilyResult(loaded, loadedFamily);
                    diagnostic.Add(
                        "doc.LoadFamily завершился: loaded = " + loaded +
                        ", loadedFamily = '" + GetFamilyNameForDiagnostic(loadedFamily) + "'.");
                }

                if (loaded && loadedFamily != null)
                    return loadedFamily;
            }
            catch (Exception ex)
            {
                if (diagnostic != null)
                    diagnostic.AddFamilyFileOpenException("Исключение во время doc.LoadFamily", ex);
            }

            if (diagnostic != null)
                diagnostic.Add("Проверяем, появилось ли семейство в проекте после doc.LoadFamily.");

            existingFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            if (!string.IsNullOrWhiteSpace(realFamilyName))
            {
                Family existingAfterLoadByRealName = existingFamilies.FirstOrDefault(f =>
                    string.Equals(f.Name, realFamilyName, StringComparison.OrdinalIgnoreCase));

                if (diagnostic != null)
                    diagnostic.Add(
                        existingAfterLoadByRealName != null
                            ? "После загрузки семейство найдено по имени из файла: '" + GetFamilyNameForDiagnostic(existingAfterLoadByRealName) + "'."
                            : "После загрузки семейство по имени из файла не найдено.");

                if (existingAfterLoadByRealName != null)
                    return existingAfterLoadByRealName;
            }

            Family existingAfterLoadByFileName = existingFamilies.FirstOrDefault(f =>
                string.Equals(f.Name, fileFamilyName, StringComparison.OrdinalIgnoreCase));

            if (diagnostic != null)
                diagnostic.Add(
                    existingAfterLoadByFileName != null
                        ? "После загрузки семейство найдено по имени файла: '" + GetFamilyNameForDiagnostic(existingAfterLoadByFileName) + "'."
                        : "После загрузки семейство по имени файла не найдено.");

            if (existingAfterLoadByFileName != null)
                return existingAfterLoadByFileName;

            if (diagnostic != null)
                diagnostic.Add("Итог: семейство не удалось загрузить или найти в проекте.");

            return null;
        }

        private static FamilySymbol GetFirstFamilySymbol(Document doc, Family family)
        {
            ISet<ElementId> symbolIds = family.GetFamilySymbolIds();
            if (symbolIds == null || symbolIds.Count == 0)
                return null;

            ElementId firstId = symbolIds.First();
            return doc.GetElement(firstId) as FamilySymbol;
        }

        private static FamilyInstance PlaceFamilyInstance(Document doc, ViewPlan floorPlan, FamilySymbol symbol, XYZ point)
        {
            FamilyPlacementType placementType = symbol.Family.FamilyPlacementType;

            switch (placementType)
            {
                case FamilyPlacementType.ViewBased:
                    return doc.Create.NewFamilyInstance(point, symbol, floorPlan);

                case FamilyPlacementType.OneLevelBased:
                case FamilyPlacementType.OneLevelBasedHosted:
                case FamilyPlacementType.WorkPlaneBased:
                    if (floorPlan.GenLevel == null)
                        throw new Exception("У активного плана не определён уровень.");

                    XYZ levelPoint = new XYZ(point.X, point.Y, 0.0);
                    FamilyInstance placed = doc.Create.NewFamilyInstance(levelPoint, symbol, floorPlan.GenLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    ResetFamilyInstanceVerticalOffsets(placed);
                    return placed;

                default:
                    throw new NotSupportedException("Тип размещения семейства не поддерживается: " + placementType);
            }
        }

        private static void ResetFamilyInstanceVerticalOffsets(FamilyInstance instance)
        {
            if (instance == null)
                return;

            TrySetDoubleParameter(instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM), 0.0);
            TrySetDoubleParameter(instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM), 0.0);

            TrySetDoubleParameterByName(instance, 0.0,
                "Отметка от уровня",
                "Elevation from Level",
                "Offset from Level");

            TrySetDoubleParameterByName(instance, 0.0,
                "Смещение от главной модели",
                "Offset from Host",
                "Offset from Main Model");
        }

        private static bool TrySetDoubleParameter(Parameter parameter, double value)
        {
            if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.Double)
                return false;

            parameter.Set(value);
            return true;
        }

        private static bool TrySetDoubleParameterByName(Element element, double value, params string[] parameterNames)
        {
            if (element == null || parameterNames == null || parameterNames.Length == 0)
                return false;

            bool changed = false;

            foreach (Parameter parameter in element.Parameters)
            {
                if (parameter == null || parameter.Definition == null)
                    continue;

                string parameterName = parameter.Definition.Name;
                if (string.IsNullOrWhiteSpace(parameterName))
                    continue;

                if (!parameterNames.Any(name => string.Equals(name, parameterName, StringComparison.OrdinalIgnoreCase)))
                    continue;

                changed |= TrySetDoubleParameter(parameter, value);
            }

            return changed;
        }

        private static void AppendComment(Element e, string textToAppend)
        {
            if (e == null || string.IsNullOrWhiteSpace(textToAppend))
                return;

            Parameter p = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (p == null || p.IsReadOnly)
                return;

            string oldValue = p.AsString();

            if (string.IsNullOrWhiteSpace(oldValue))
            {
                p.Set(textToAppend);
                return;
            }

            if (oldValue.Contains(textToAppend))
                return;

            p.Set(oldValue + " " + textToAppend);
        }
    }
}