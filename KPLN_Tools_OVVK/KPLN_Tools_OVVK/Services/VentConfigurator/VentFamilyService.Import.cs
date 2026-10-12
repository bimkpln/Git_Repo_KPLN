using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal sealed partial class VentFamilyService
    {
        private bool EnsureProjectFamily(string path)
        {
            bool ready = ProjectFamilyStorage.EnsureFile(ResolveSourcePath(), path, (source, destination) =>
            {
                var dialog = new TaskDialog(PluginName)
                {
                    MainInstruction = "Семейство проекта не найдено. Создать его?",
                    MainContent = "Файл проекта:\n" + destination
                        + "\n\nБудет создана папка, если её нет, и скопировано исходное семейство:\n" + source,
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    DefaultButton = TaskDialogResult.No
                };
                return dialog.Show() == TaskDialogResult.Yes;
            });
            if (!ready)
                _owner.SetStatus("Создание проектного семейства отменено.", false);
            return ready;
        }

        private static Family FindLoadedInstallationFamily(Document document, string projectCode)
        {
            if (document == null || !document.IsValidObject || document.IsFamilyDocument)
                return null;
            string expected = Path.GetFileNameWithoutExtension(VentConfiguratorSettings.MakeFileName(projectCode));
            var families = new FilteredElementCollector(document).OfClass(typeof(Family)).Cast<Family>()
                .Where(f => f.IsEditable && (f.Name.StartsWith("550_Универсальная установка_Одноуровневая_(", StringComparison.OrdinalIgnoreCase)
                    || f.Name.StartsWith("Вентустановка_", StringComparison.OrdinalIgnoreCase))).ToList();
            var exact = families.Where(f => string.Equals(f.Name, expected, StringComparison.OrdinalIgnoreCase)).ToList();
            if (exact.Count == 1)
                return exact[0];
            if (families.Count > 1)
                throw new InvalidOperationException("В проекте несколько семейств вентустановок. Для однозначной загрузки нужно задать папку проектного семейства.");
            return families.SingleOrDefault();
        }

        // RVT предоставляет значения типов напрямую. EditFamily здесь не нужен:
        // его геометрический решатель может отказать ещё до чтения параметров.
        private sealed class ProjectTypeImport
        {
            internal string Name;
            internal Dictionary<string, object> Values = new Dictionary<string, object>(StringComparer.Ordinal);
            internal Dictionary<string, string> ReadErrors = new Dictionary<string, string>(StringComparer.Ordinal);
        }

        private sealed class PendingProjectImport
        {
            internal string SnapshotPath, DestinationPath, DestinationHash;
            internal bool FromLibrary;

            internal void CheckDestination(string path)
            {
                bool same = SamePath(path, DestinationPath);
                if (!same && File.Exists(path))
                    throw new InvalidOperationException("Для перенесённых типов выберите новый файл либо исходный выбранный RFA. Другой существующий файл может содержать свои типы.");
                if (same && DestinationHash != (File.Exists(path) ? FileHash(path) : null))
                    throw new IOException("Выбранный RFA изменился после чтения. Он не перезаписан; заново откройте его для переноса типов.");
            }
        }

        private FamilyPackage ReadSelectedFamilyFile(UIApplication app, Document project, Family family, string failure)
        {
            if (_owner.HasUnsavedEdits)
                throw new InvalidOperationException("Перед сменой файла сохраните изменённые типы.");
            string path = _owner.ChooseExistingFamilyPath(family?.Name, failure);
            if (path == null)
                throw new System.OperationCanceledException();
            if (family != null)
                return ImportProjectFamily(app, project, family, path);
            var opened = FindOpenFamily(app, path);
            var package = opened == null ? ReadFamilyPackage(app, path, true) : ReadOpenFamilyPackage(app, opened, path);
            if (_activeContextKey != null)
                _projectImports.Remove(_activeContextKey);
            if (_activeContextKey != null && package.Path != null)
                SavedFamilyPaths[_activeContextKey] = package.Path;
            return package;
        }

        private FamilyPackage ReadProjectFamily(UIApplication app, Document project, Family family)
        {
            if (project.IsModifiable)
                throw new InvalidOperationException("Завершите текущую команду Revit перед чтением семейства.");
            string path = ResolveSourcePath();
            if (!File.Exists(path))
                return ReadSelectedFamilyFile(app, project, family, "Исходное семейство недоступно: " + path);
            return ImportProjectFamily(app, project, family, path);
        }

        private sealed class ProjectParameterIndex
        {
            internal readonly ILookup<string, Parameter> ByName;
            internal readonly ILookup<Guid, Parameter> ByGuid;

            internal ProjectParameterIndex(Element element)
            {
                // Один обход определений на тип; значения вне интерфейса не читаем.
                var parameters = element.Parameters.Cast<Parameter>().ToList();
                ByName = parameters.ToLookup(p => FamilyParameterNames.Canonical(p.Definition.Name), StringComparer.Ordinal);
                ByGuid = parameters.Where(p => p.IsShared).ToLookup(p => p.GUID);
            }
        }

        private static Parameter FindProjectParameter(ProjectParameterIndex parameters, FamilyParameter target)
        {
            if (target.IsShared)
            {
                var shared = parameters.ByGuid[target.GUID].ToList();
                if (shared.Count > 1)
                    throw new InvalidOperationException("Несколько параметров с GUID " + target.GUID + ".");
                if (shared.Count == 1)
                    return shared[0];
            }
            var matches = parameters.ByName[FamilyParameterNames.Canonical(target.Definition.Name)].ToList();
            if (matches.Count > 1)
                throw new InvalidOperationException("Неоднозначный параметр «" + target.Definition.Name + "» в модели.");
            var result = matches.SingleOrDefault();
            if (result != null && result.IsShared && target.IsShared && result.GUID != target.GUID)
                throw new InvalidOperationException("У одноимённого общего параметра в RVT и RFA разные GUID.");
            return result;
        }

        private static object ReadProjectInput(Document project, Parameter source, FamilyParameter target)
        {
            if (!Common.IDHelper.HaveCompatibleParameterDataTypes(source, target))
                throw new InvalidOperationException("Тип данных параметра в модели отличается от основы. Введите значение вручную"
                    + (target.StorageType == StorageType.Double ? " в миллиметрах" : "") + " или используйте значение по умолчанию конфигуратора.");
            switch (target.StorageType)
            {
                case StorageType.String:
                    return source.AsString() ?? string.Empty;
                case StorageType.Integer:
                    return source.AsInteger();
                case StorageType.Double:
                    return Common.IDHelper.ConvertInternalToMm(source.AsDouble());
                case StorageType.ElementId:
                    var id = source.AsElementId();
                    if (id == null || id.Equals(ElementId.InvalidElementId))
                        throw new InvalidOperationException("В модели не выбран тип элемента. Выберите его в составе установки.");
                    return DescribeType(project, id); // только имена; никаких ID между документами
                default:
                    throw new InvalidOperationException("Значение не прочитано. Задайте его в конфигураторе.");
            }
        }

        private static List<ProjectTypeImport> CaptureProjectTypeInputs(Document project, Family family, Document destination)
        {
            var manager = destination.FamilyManager;
            var result = new List<ProjectTypeImport>();
            var targets = new Dictionary<string, FamilyParameter>(StringComparer.Ordinal);
            var targetErrors = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string name in InstallationConfiguration.InterfaceParameterNames().Select(FamilyParameterNames.Canonical).Distinct())
            {
                try
                {
                    targets.Add(name, FindFamilyParameter(manager, name));
                }
                catch (Exception ex) { targetErrors.Add(name, OperationError.Message(ex)); }
            }
            var infoNames = new HashSet<string>(new InstallationInfo().Assignments().Keys, StringComparer.Ordinal);
            foreach (var id in family.GetFamilySymbolIds())
            {
                var symbol = project.GetElement(id) as FamilySymbol;
                if (symbol == null)
                    continue;
                if (symbol.Name == BaseFamilyTypeName)
                    continue;
                var type = new ProjectTypeImport { Name = symbol.Name, ReadErrors = new Dictionary<string, string>(targetErrors, StringComparer.Ordinal) };
                var parameters = new ProjectParameterIndex(symbol);
                // Белый список интерфейса. Параметры материалов, производительности,
                // служебные флаги, ограничения и формулы из RVT здесь вообще не читаются.
                foreach (var pair in targets)
                {
                    string name = pair.Key;
                    try
                    {
                        var target = pair.Value;
                        if (target == null)
                        {
                            type.ReadErrors[name] = "В основе отсутствует параметр. Обновите основу семейства.";
                            continue;
                        }
                        bool valve = name == ValveControl.ParameterName && ValveControl.ConstantFormula(target.Formula).HasValue;
                        bool info = infoNames.Contains(name);
                        if (!info && (target.IsReadOnly || target.IsReporting || target.IsDeterminedByFormula && !valve))
                            continue;
                        // Экземпляры одного типа вправе иметь разные значения. Они не задают
                        // настройки типа: недостающее берём из одноимённого RFA или значений по умолчанию.
                        var source = FindProjectParameter(parameters, target);
                        FamilyParameter infoValue, infoMode;
                        if (info && TryGetSystemNameInputs(manager, target, out infoValue, out infoMode))
                        {
                            // У формульного результата может не быть значения на FamilySymbol.
                            // Читаем именно вход типа; в старом RVT возможен только прямой КП_О.
                            var typeSource = FindProjectParameter(parameters, infoValue);
                            if (typeSource != null && typeSource.HasValue)
                            {
                                source = typeSource;
                                target = infoValue;
                            }
                        }
                        type.Values[name] = source == null || !source.HasValue ? null : ReadProjectInput(project, source, target);
                    }
                    catch (Exception ex) { type.ReadErrors[name] = OperationError.Message(ex); }
                }
                result.Add(type);
            }
            return result;
        }

        private FamilyPackage ImportProjectFamily(UIApplication app, Document project, Family family, string path)
        {
            ValidateDiskSource(app, path);
            string originalHash = FileHash(path);
            bool library = SamePath(path, ResolveSourcePath()) || SamePath(path, SourceFamilyPath) || SamePath(path, LiteralSourceFamilyPath);
            string destinationPath = library ? null : path;
            string directory = Path.Combine(Path.GetTempPath(), "KPLN_Ventilation_Project_" + Guid.NewGuid().ToString("N"));
            string input = Path.Combine(directory, Path.GetFileName(path));
            Document document = null;
            bool keep = false;
            try
            {
                Directory.CreateDirectory(directory);
                File.Copy(path, input, false);
                File.SetAttributes(input, File.GetAttributes(input) & ~FileAttributes.ReadOnly);
                if (FileHash(input) != originalHash)
                    throw new IOException("Основа изменилась во время чтения. Повторите открытие.");
                document = app.Application.OpenDocumentFile(input);
                if (!document.IsFamilyDocument)
                    throw new InvalidOperationException("Выбранный файл не является семейством.");
                var package = ReadFamilyPackage(document, destinationPath, !library);
                var types = CaptureProjectTypeInputs(project, family, document);
                foreach (var type in types)
                {
                    var item = package.Types.SingleOrDefault(t => t.Name == type.Name);
                    if (item == null)
                    {
                        item = new FamilyTypeItem { Name = type.Name, SavedPath = destinationPath };
                        package.Types.Add(item); // PersistedName пуст: пока это настройки, а не созданный RFA-тип.
                    }
                    item.LoadError = null;
                    var savedConfiguration = item.Configuration;
                    item.Configuration = InstallationConfiguration.FromProjectInputs(package.Catalog, type.Values, type.ReadErrors, savedConfiguration);
                    item.HasAutomaticName = false;
                    if (savedConfiguration != null && savedConfiguration.SameInputs(item.Configuration) && item.Configuration.SameInputs(savedConfiguration)
                        && item.Configuration.Problems().Count == 0)
                        item.MarkSaved();
                    else
                        item.MarkImported();
                }
                if (!document.Close(false))
                    throw new InvalidOperationException("Не удалось закрыть копию основы.");
                document = null;
                if (FileHash(path) != originalHash)
                    throw new IOException("Основа изменилась во время чтения. Повторите открытие.");
                package.SourcePath = input;
                foreach (var item in package.Types)
                    item.SourcePath = input;
                int problems = package.Types.Count(t => t.Configuration != null && t.Configuration.Problems().Count > 0);
                package.LoadNotice = "Загружено типов: " + package.Types.Count + ". Сохраняется только выбранный тип."
                    + (problems == 0 ? "" : " Есть незаполненные или некорректные поля у типов: " + problems + ". Выберите тип и исправьте отмеченные поля.");
                _projectImports[_activeContextKey] = new PendingProjectImport
                {
                    SnapshotPath = input,
                    DestinationPath = destinationPath,
                    DestinationHash = library ? null : originalHash,
                    FromLibrary = library
                };
                _automaticReadDirectories.Add(directory);
                keep = true;
                return package;
            }
            finally
            {
                bool closed = CloseCompositionDocument(document);
                if (!keep && closed)
                    DeleteCompositionDirectory(directory);
            }
        }

        internal void ClearAutomaticSnapshots()
        {
            foreach (string directory in _automaticReadDirectories)
            {
                try
                {
                    if (Directory.Exists(directory))
                        Directory.Delete(directory, true);
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
            _automaticReadDirectories.Clear();
        }
    }
}
