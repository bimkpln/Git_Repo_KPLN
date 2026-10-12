using System;
using System.Collections.Generic;
using System.Globalization;
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
        private static CompositionSnapshot ReadCompositionSnapshot(Document document, bool includeSettings)
        {
            var result = new CompositionSnapshot();
            for (int slot = 1; slot <= 12; slot++)
                result.Choices.Add(slot, new HashSet<string>(ReadTypeChoices(document, SectionParameter(document.FamilyManager, slot))
                    .Select(pair => pair.Value.Key), StringComparer.Ordinal));
            if (!includeSettings)
                return result;
            var inputs = new HashSet<string>(InstallationConfiguration.InterfaceParameterNames().Select(FamilyParameterNames.Canonical), StringComparer.Ordinal);
            var parameters = document.FamilyManager.Parameters.Cast<FamilyParameter>().ToList();
            foreach (var parameter in parameters)
                result.Settings.Add("Параметр\n" + parameter.Definition.Name, parameter.StorageType + "|" + parameter.IsInstance
                    + "|" + parameter.IsReporting + "|" + parameter.IsReadOnly + "|" + (parameter.Formula ?? string.Empty));
            foreach (FamilyType type in document.FamilyManager.Types)
            {
                result.Settings.Add("Тип\n" + type.Name, true);
                foreach (var parameter in parameters)
                {
                    if (parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting)
                        continue;
                    // Служебная геометрия может штатно пересчитаться от новых вложенных определений.
                    if (parameter.StorageType == StorageType.Double && !inputs.Contains(FamilyParameterNames.Canonical(parameter.Definition.Name)))
                        continue;
                    object value = ReadTypeInput(type, parameter);
                    if (parameter.StorageType == StorageType.ElementId)
                    {
                        var id = type.AsElementId(parameter);
                        var element = id == null ? null : document.GetElement(id);
                        value = element is ElementType || element is NestedFamilyTypeReference ? DescribeType(document, id).Key
                            : element == null ? value : element.GetType().FullName + "\n" + element.Name;
                    }
                    result.Settings.Add(type.Name + "\n" + parameter.Definition.Name, value);
                }
            }
            return result;
        }

        private static Family ReloadNestedFamily(UIApplication app, Document nested, Document target, IFamilyLoadOptions options)
        {
            var failures = new TransactionFailures(true);
            bool failed = false;
            EventHandler<Autodesk.Revit.DB.Events.FailuresProcessingEventArgs> handler = (sender, args) =>
            {
                var accessor = args.GetFailuresAccessor();
                if (!accessor.GetDocument().Equals(target))
                    return;
                var result = failures.PreprocessFailures(accessor);
                if (result == FailureProcessingResult.ProceedWithRollBack)
                {
                    failed = true;
                    var handling = accessor.GetFailureHandlingOptions();
                    handling.SetClearAfterRollback(true);
                    accessor.SetFailureHandlingOptions(handling);
                    args.SetProcessingResult(result);
                }
            };
            app.Application.FailuresProcessing += handler;
            try
            {
                // Эта перегрузка сама ведёт транзакцию; целевой документ не должен быть IsModifiable.
                // Revit вызывает IFamilyLoadOptions только для реально изменённых загруженных определений.
                var loaded = nested.LoadFamily(target, options);
                if (failed)
                    throw new InvalidOperationException(failures.DescribeErrors("Revit отменил обновление вложенного семейства."));
                return loaded;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(failures.DescribeErrors("Не удалось загрузить семейство: " + OperationError.Message(ex)), ex);
            }
            finally { app.Application.FailuresProcessing -= handler; }
        }

        private static Document OpenCompositionCopy(UIApplication app, string source, string directory)
        {
            Directory.CreateDirectory(directory);
            string copy = Path.Combine(directory, Path.GetFileName(source));
            File.Copy(source, copy, false);
            File.SetAttributes(copy, File.GetAttributes(copy) & ~FileAttributes.ReadOnly);
            var document = app.Application.OpenDocumentFile(copy);
            if (document.IsFamilyDocument)
                return document;
            document.Close(false);
            throw new InvalidOperationException("Файл состава не является семейством Revit: " + source);
        }

        private static bool CloseCompositionDocument(Document document)
        {
            if (document == null || !document.IsValidObject)
                return true;
            try
            {
                return document.Close(false);
            }
            catch { return false; }
        }

        private static void DeleteCompositionDirectory(string directory)
        {
            if (directory == null)
                return;
            try
            {
                if (Directory.Exists(directory))
                    Directory.Delete(directory, true);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        // Проверка при запуске только читает определения. Не загружаем семейства
        // друг в друга даже во временных документах: это запускало пересчёт Revit.
        private static CompositionSnapshot ReadFamilyDefinitionSnapshot(Document document, SectionCatalog catalog = null)
        {
            var result = new CompositionSnapshot();
            if (catalog == null)
                result = ReadCompositionSnapshot(document, false);
            else
                foreach (var slot in catalog)
                    result.Choices.Add(slot.Key, new HashSet<string>(slot.Value.Choices.Select(c => c.Key), StringComparer.Ordinal));
            foreach (FamilyParameter parameter in document.FamilyManager.Parameters)
            {
                string name = FamilyParameterNames.Canonical(parameter.Definition.Name);
                string formula = parameter.Formula ?? string.Empty;
                // Состояние клапана — настройка типа, а не изменение основы.
                if (name == ValveControl.ParameterName && ValveControl.ConstantFormula(formula).HasValue)
                    formula = string.Empty;
                result.Settings.Add(name, parameter.StorageType + "|" + parameter.IsInstance
                    + "|" + Common.IDHelper.ParameterDataType(parameter) + "|" + (parameter.IsShared ? parameter.GUID.ToString() : string.Empty)
                    + "|" + parameter.IsReadOnly + "|" + parameter.IsReporting + "|" + formula);
            }
            return result;
        }

        private static bool FamilyDefinitionDiffers(CompositionSnapshot source, CompositionSnapshot project)
        {
            if (source.Settings.Count != project.Settings.Count || source.Settings.Any(p =>
                !project.Settings.ContainsKey(p.Key) || !Equals(project.Settings[p.Key], p.Value)))
                return true;
            // Дополнительные типы проекта допустимы и после обновления сохраняются.
            return source.Choices.Any(p => !project.Choices.ContainsKey(p.Key) || !p.Value.IsSubsetOf(project.Choices[p.Key]));
        }

        private string CheckProjectFamilyUpdates(UIApplication app, string projectPath, CompositionSnapshot projectDefinition = null)
        {
            if (string.IsNullOrWhiteSpace(projectPath))
                return null;
            string directory = null;
            Document source = null, target = null;
            try
            {
                string sourcePath = ResolveSourcePath();
                if (SamePath(projectPath, sourcePath))
                    return null;
                if (!File.Exists(sourcePath))
                    return "Основа семейства недоступна. Проверка обновлений не выполнена. Обратитесь к BIM-координатору.";
                string sourceHash = FileHash(sourcePath), projectHash = FileHash(projectPath);
                if (sourceHash == projectHash)
                    return null;
                directory = Path.Combine(Path.GetTempPath(), "KPLN_FamilyCheck_" + Guid.NewGuid().ToString("N"));
                source = OpenCompositionCopy(app, sourcePath, Path.Combine(directory, "source"));
                if (projectDefinition == null)
                {
                    target = OpenCompositionCopy(app, projectPath, Path.Combine(directory, "project"));
                    projectDefinition = ReadFamilyDefinitionSnapshot(target);
                }
                bool differs = FamilyDefinitionDiffers(ReadFamilyDefinitionSnapshot(source), projectDefinition);
                if (sourceHash != FileHash(sourcePath) || projectHash != FileHash(projectPath))
                    return "Семейство изменилось во время проверки обновлений. Повторите проверку при следующем открытии конфигуратора.";
                if (!differs)
                    return null;
                return CanUpdateProjectFamily()
                    ? "Семейство проекта отличается от основы. Для обновления нажмите значок рядом с «Типы семейства»."
                    : "Семейство проекта отличается от основы. Попросите вашего BIM-координатора обновить семейство.";
            }
            catch
            {
                // Проверка обновлений не должна блокировать чтение сохранённых типов.
                return "Не удалось проверить обновления семейства. Обратитесь к BIM-координатору. Используется сохранённое семейство проекта.";
            }
            finally
            {
                bool targetClosed = CloseCompositionDocument(target);
                bool sourceClosed = CloseCompositionDocument(source);
                if (targetClosed && sourceClosed)
                    DeleteCompositionDirectory(directory);
            }
        }

        private sealed class FullFamilyReloadOptions : IFamilyLoadOptions
        {
            private readonly string _familyName;

            internal FullFamilyReloadOptions(string familyName)
            {
                _familyName = familyName;
            }

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = false;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = !string.Equals(sharedFamily.Name, _familyName, StringComparison.OrdinalIgnoreCase);
                return true;
            }
        }

        private static FamilyUpgradeSnapshot ReadUpgradeSnapshot(Document document)
        {
            var result = new FamilyUpgradeSnapshot();
            var manager = document.FamilyManager;
            var interfaceNames = new HashSet<string>(InstallationConfiguration.InterfaceParameterNames().Select(FamilyParameterNames.Canonical));
            var parameters = manager.Parameters.Cast<FamilyParameter>().ToList();
            foreach (var p in parameters)
                result.Parameters.Add(FamilyParameterNames.Canonical(p.Definition.Name), new FamilyUpgradeParameter
                {
                    Kind = p.StorageType + "|" + p.IsInstance + "|" + Common.IDHelper.ParameterDataType(p)
                        + "|" + (p.IsShared ? p.GUID.ToString() : string.Empty),
                    Formula = p.Formula ?? string.Empty,
                    ReadOnly = p.IsReadOnly,
                    Reporting = p.IsReporting
                });
            foreach (FamilyType type in manager.Types)
            {
                var values = new Dictionary<string, object>(StringComparer.Ordinal);
                result.Types.Add(type.Name, values);
                foreach (var p in parameters)
                {
                    string name = FamilyParameterNames.Canonical(p.Definition.Name);
                    bool valve = name == ValveControl.ParameterName && ValveControl.ConstantFormula(p.Formula).HasValue;
                    if ((!valve && p.IsDeterminedByFormula) || p.IsReadOnly || p.IsReporting)
                        continue;
                    // Служебные геометрические размеры вправе пересчитываться при обновлении основы.
                    if (p.StorageType == StorageType.Double && !interfaceNames.Contains(name))
                        continue;
                    object value = ReadTypeInput(type, p);
                    if (p.StorageType == StorageType.ElementId)
                    {
                        var id = type.AsElementId(p);
                        var element = id == null ? null : document.GetElement(id);
                        value = element is ElementType || element is NestedFamilyTypeReference ? DescribeType(document, id).Key
                            : element == null ? value : element.GetType().FullName + "\n" + element.Name;
                    }
                    values.Add(name, value);
                }
            }
            return result;
        }

        private static void RestoreUpgradeValves(Document document, FamilyUpgradeSnapshot before, Action<string> setStage)
        {
            var manager = document.FamilyManager;
            var parameter = FindFamilyParameter(manager, ValveControl.ParameterName);
            if (parameter == null || !ValveControl.CanEdit(new ParameterDependency
            {
                Formula = parameter.Formula,
                IsReadOnly = parameter.IsReadOnly,
                IsReporting = parameter.IsReporting
            }))
                return;
            RunFamilyStage(document, "Сохранение состояния клапанов типов", () =>
            {
                foreach (var item in before.Types)
                {
                    object value;
                    if (!item.Value.TryGetValue(ValveControl.ParameterName, out value) || value == null)
                        continue;
                    SelectFamilyType(manager, item.Key);
                    ApplyValveState(manager, Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0);
                }
            }, null, setStage);
        }

        private void EnsureUpgradeUnchanged(UIApplication app, string projectPath, string sourcePath, string projectHash, string sourceHash)
        {
            if (!CanUpdateProjectFamily())
                throw new UnauthorizedAccessException("Обновление семейства доступно только отделу 8.");
            if (FindOpenFamily(app, projectPath) != null)
                throw new InvalidOperationException("Закройте семейство проекта в Revit перед его обновлением.");
            ValidateDiskSource(app, sourcePath);
            if (FileHash(projectPath) != projectHash || FileHash(sourcePath) != sourceHash)
                throw new IOException("Семейство изменилось во время подготовки обновления. Файл проекта не заменён; повторите действие.");
        }

        private static void RemoveSourceExampleTypes(Document document, Action<string> setStage)
        {
            RunFamilyStage(document, "Подготовка новой основы", () =>
            {
                var manager = document.FamilyManager;
                foreach (string name in manager.Types.Cast<FamilyType>().Select(t => t.Name).Where(n => n != BaseFamilyTypeName).ToList())
                    DeleteFamilyType(manager, name);
                SelectFamilyType(manager, BaseFamilyTypeName);
            }, null, setStage);
        }

        private static List<FamilyTypeItem> CollectUpdateTypes(FamilyPackage saved, IList<FamilyTypeItem> edited)
        {
            var result = saved.Types.ToList();
            foreach (var item in edited ?? new List<FamilyTypeItem>())
            {
                var original = result.FirstOrDefault(t => t.Name == (item.PersistedName ?? item.Name));
                if (original != null && !item.IsDirty && original.Configuration != null)
                    continue;
                if (original != null)
                    result.Remove(original);
                if (result.Any(t => string.Equals(t.Name, item.Name, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Повторяется имя типа «" + item.Name + "». Исправьте имя перед обновлением.");
                result.Add(item);
            }
            return result;
        }

        private static void VerifyRebuiltTypes(FamilyPackage package, IDictionary<string, InstallationConfiguration> expected)
        {
            if (package.Types.Count != expected.Count || package.Types.Any(t => !expected.ContainsKey(t.Name)))
                throw new InvalidOperationException("После обновления изменился список типов проекта. Файл проекта не заменён.");
            foreach (var pair in expected)
            {
                var actual = package.Types.SingleOrDefault(t => t.Name == pair.Key)?.Configuration;
                if (actual == null || !pair.Value.SameInputs(actual))
                    throw new InvalidOperationException("Не сохранены настройки типа «" + pair.Key + "». Файл проекта не заменён.");
            }
        }

        private FamilyPackage UpdateEntireProjectFamily(UIApplication app, string projectPath, string selectedType, IList<FamilyTypeItem> editedTypes, Action<string> setStage, out string backupPath)
        {
            backupPath = null;
            if (!CanUpdateProjectFamily())
                throw new UnauthorizedAccessException("Обновление семейства доступно только отделу 8.");
            string sourcePath = ResolveSourcePath();
            if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
                throw new InvalidOperationException("Сначала сохраните семейство проекта.");
            if (SamePath(projectPath, sourcePath) || SamePath(projectPath, SourceFamilyPath) || SamePath(projectPath, LiteralSourceFamilyPath))
                throw new InvalidOperationException("Обновлять можно только проектную копию, а не исходное семейство на X.");
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("Исходное семейство на X недоступно.", sourcePath);
            string projectHash = FileHash(projectPath), sourceHash = FileHash(sourcePath);
            EnsureUpgradeUnchanged(app, projectPath, sourcePath, projectHash, sourceHash);
            string directory = Path.Combine(Path.GetTempPath(), "KPLN_FamilyUpdate_" + Guid.NewGuid().ToString("N"));
            string staging = null;
            Document previous = null, updated = null;
            try
            {
                setStage("Чтение настроек текущих типов проекта");
                previous = OpenCompositionCopy(app, projectPath, Path.Combine(directory, "previous"));
                var types = CollectUpdateTypes(ReadFamilyPackage(previous, projectPath, true), editedTypes);
                if (!previous.Close(false))
                    throw new InvalidOperationException("Не удалось закрыть копию текущего семейства.");
                previous = null;
                updated = OpenCompositionCopy(app, sourcePath, Path.Combine(directory, "new"));
                var catalog = ReadFamilyPackage(updated, null, false).Catalog;
                var expected = new Dictionary<string, InstallationConfiguration>(StringComparer.Ordinal);
                foreach (var item in types)
                {
                    string name = VentTypeName.ValidateTypeName(item.Name);
                    if (item.Configuration == null)
                        throw new InvalidOperationException(item.LoadError ?? "Не прочитаны настройки типа «" + name + "».");
                    var unresolved = item.Configuration.Problems().Where(p => item.Configuration.ReviewIssues.Contains(p)).ToList();
                    // Сначала переносим поля на новую схему: отсутствующие значения получают
                    // значения по умолчанию, а реальные ошибки остаются у соответствующего поля.
                    var configuration = InstallationConfiguration.FromProjectInputs(catalog, item.Configuration.EditableInputs(),
                        unresolved.ToDictionary(p => p.ParameterName, p => p.Message, StringComparer.Ordinal));
                    try
                    {
                        configuration.ValidateForOutput();
                    }
                    catch (Exception ex) { throw new InvalidOperationException("Тип «" + name + "»: " + OperationError.Message(ex), ex); }
                    if (expected.ContainsKey(name))
                        throw new InvalidOperationException("Повторяется имя типа «" + name + "».");
                    expected.Add(name, configuration);
                }
                // Заполненные примеры типов на X не добавляем в проект. Шаблон сохраняем
                // для создания новых типов; его геометрию отдельно не перестраиваем.
                RemoveSourceExampleTypes(updated, setStage);
                foreach (var pair in expected)
                {
                    string typeName = pair.Key;
                    ApplyFamilyMutation(updated, sourcePath, typeName, null, true, pair.Value, false,
                        value => setStage("Обновление типа «" + typeName + "»: " + value), value => { });
                }
                var package = ReadFamilyPackage(updated, projectPath, true);
                VerifyRebuiltTypes(package, expected);
                var dialog = new TaskDialog(PluginName)
                {
                    MainInstruction = "Обновить семейство проекта из основы на X?",
                    MainContent = "На новой основе подготовлено типов: " + expected.Count + ".\n"
                        + "Сохранены имена, состав и настройки конфигуратора ваших типов, включая текущие изменения в интерфейсе.\n"
                        + "Заполненные значения и примеры типов из X в настройки проекта не перенесены. Новые поля получили значения по умолчанию конфигуратора.\n"
                        + "Рядом с файлом будет сохранена резервная копия.\n\n" + projectPath,
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    DefaultButton = TaskDialogResult.No
                };
                if (dialog.Show() != TaskDialogResult.Yes)
                    return null;
                EnsureUpgradeUnchanged(app, projectPath, sourcePath, projectHash, sourceHash);
                staging = Path.Combine(Path.GetDirectoryName(projectPath), ".KPLN_FamilyUpdate_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(staging);
                string staged = Path.Combine(staging, Path.GetFileName(projectPath));
                setStage("Сохранение и проверка обновлённого семейства");
                using (var save = new SaveAsOptions { OverwriteExistingFile = false, MaximumBackups = 1 })
                    updated.SaveAs(staged, save);
                if (!updated.Close(false))
                    throw new InvalidOperationException("Не удалось закрыть подготовленное семейство.");
                updated = app.Application.OpenDocumentFile(staged);
                package = ReadFamilyPackage(updated, projectPath, true);
                VerifyRebuiltTypes(package, expected);
                package.SelectedTypeName = selectedType;
                if (!updated.Close(false))
                    throw new InvalidOperationException("Не удалось закрыть проверенное семейство.");
                updated = null;
                EnsureUpgradeUnchanged(app, projectPath, sourcePath, projectHash, sourceHash);
                string backup = Path.Combine(Path.GetDirectoryName(projectPath), Path.GetFileNameWithoutExtension(projectPath)
                    + "_до_обновления_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".rfa");
                setStage("Замена семейства проекта");
                File.Replace(staged, projectPath, backup);
                backupPath = backup;
                return package;
            }
            finally
            {
                bool updatedClosed = CloseCompositionDocument(updated), previousClosed = CloseCompositionDocument(previous);
                if (updatedClosed && previousClosed)
                    DeleteCompositionDirectory(directory);
                if (updatedClosed)
                    DeleteCompositionDirectory(staging);
            }
        }
    }
}
