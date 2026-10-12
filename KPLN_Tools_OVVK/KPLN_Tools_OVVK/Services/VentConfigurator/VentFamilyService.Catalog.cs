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
        private static string DependencyElement(Element element)
        {
            if (element == null)
                return "Элемент не найден";
            var instance = element as FamilyInstance;
            string name;
            try
            {
                name = instance?.Symbol == null ? element.Name : instance.Symbol.FamilyName + " : " + instance.Symbol.Name;
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException) { name = element.GetType().Name; }
            return name + " [ID " + Common.IDHelper.ElIdValue(element.Id) + "]";
        }

        private static FamilyDependencies ReadCatalogDependencies(Document document, FamilyType sourceType, FamilyDependencies schema = null)
        {
            var result = new FamilyDependencies();
            if (schema == null)
            {
                // Для интерфейса нужны определения и формулы. Обход привязок геометрии
                // и вложенных элементов выполняется только в диагностике операции записи.
                foreach (FamilyParameter parameter in document.FamilyManager.Parameters)
                    result.Parameters.Add(parameter.Definition.Name, new ParameterDependency
                    {
                        Name = parameter.Definition.Name,
                        Formula = parameter.Formula,
                        IsReadOnly = parameter.IsReadOnly,
                        IsReporting = parameter.IsReporting,
                        IsInstance = parameter.IsInstance
                    });
                result.ConnectFormulas();
            }
            else
                result.Parameters = new Dictionary<string, ParameterDependency>(schema.Parameters, StringComparer.Ordinal);
            // Только эти четыре флага участвуют в локальном расчёте. Их значения
            // принадлежат конкретному типу; общий снимок определений не изменяем.
            foreach (string name in new[] { "Соединитель_Приточный", "Соединитель_Отработанный", "Системный_Видимость_1_Блок", ValveControl.ParameterName })
            {
                ParameterDependency definition;
                if (!result.Parameters.TryGetValue(name, out definition))
                    continue;
                var parameter = document.FamilyManager.get_Parameter(name);
                var item = new ParameterDependency
                {
                    Name = definition.Name,
                    Formula = definition.Formula,
                    IsReadOnly = definition.IsReadOnly,
                    IsReporting = definition.IsReporting,
                    IsInstance = definition.IsInstance,
                    Inputs = definition.Inputs,
                    Dependents = definition.Dependents
                };
                if (parameter != null && parameter.StorageType == StorageType.Integer)
                    item.IntegerValue = sourceType.AsInteger(parameter);
                result.Parameters[name] = item;
            }
            return result;
        }

        private static FamilyDependencies ReadFamilyDependencies(Document document, FamilyType sourceType = null, IEnumerable<string> watchedNames = null)
        {
            var result = new FamilyDependencies();
            var parameters = document.FamilyManager.Parameters.Cast<FamilyParameter>().ToList();
            sourceType = sourceType ?? document.FamilyManager.Types.Cast<FamilyType>().FirstOrDefault(t => t.Name == BaseFamilyTypeName)
                ?? document.FamilyManager.CurrentType;
            var byElement = new Dictionary<long, HashSet<string>>();
            var editable = new HashSet<string>((watchedNames ?? InstallationConfiguration.InterfaceParameterNames()).Select(FamilyParameterNames.Canonical), StringComparer.Ordinal);
            foreach (var parameter in parameters)
            {
                var item = new ParameterDependency
                {
                    Name = parameter.Definition.Name,
                    Formula = parameter.Formula,
                    IsReadOnly = parameter.IsReadOnly,
                    IsReporting = parameter.IsReporting,
                    IsInstance = parameter.IsInstance
                };
                if (sourceType != null)
                {
                    try
                    {
                        item.SourceValue = parameter.StorageType == StorageType.String ? sourceType.AsString(parameter) : sourceType.AsValueString(parameter);
                        if (parameter.StorageType == StorageType.Integer)
                            item.IntegerValue = sourceType.AsInteger(parameter);
                        if (parameter.StorageType == StorageType.ElementId)
                        {
                            var id = sourceType.AsElementId(parameter);
                            item.SourceValue = id == null ? "не задано" : DependencyElement(document.GetElement(id));
                        }
                    }
                    catch (Exception ex) { item.SourceValue = "не прочитано: " + OperationError.Message(ex); }
                }
                result.Parameters.Add(item.Name, item);
            }
            result.ConnectFormulas();
            foreach (var parameter in parameters)
            {
                string name = parameter.Definition.Name;
                try
                {
                    foreach (Parameter associated in parameter.AssociatedParameters)
                    {
                        var element = associated.Element;
                        result.Parameters[name].Links.Add("Связь: " + DependencyElement(element) + " → " + associated.Definition.Name);
                        if (element == null || !editable.Contains(FamilyParameterNames.Canonical(name)))
                            continue;
                        long id = Common.IDHelper.ElIdValue(element.Id);
                        HashSet<string> names;
                        if (!byElement.TryGetValue(id, out names))
                            byElement[id] = names = new HashSet<string>();
                        names.Add(name);
                    }
                }
                catch (Exception ex) { result.ReadWarnings.Add("Не удалось полностью прочитать связи «" + name + "»: " + OperationError.Message(ex)); }
            }
            foreach (Dimension dimension in new FilteredElementCollector(document).OfClass(typeof(Dimension)))
            {
                try
                {
                    var related = new HashSet<string>();
                    string labelStatus = "без метки";
                    // У выравниваний FamilyLabel недоступен: это не повод потерять сами привязки.
                    try
                    {
                        var label = dimension.FamilyLabel;
                        if (label != null)
                        {
                            related.Add(label.Definition.Name);
                            labelStatus = label.Definition.Name;
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException) { labelStatus = "метка неприменима"; }
                    var references = dimension.References == null ? new List<Reference>() : dimension.References.Cast<Reference>().ToList();
                    foreach (long id in references.Select(r => Common.IDHelper.ElIdValue(r.ElementId)).Distinct())
                    {
                        HashSet<string> names;
                        if (byElement.TryGetValue(id, out names))
                            related.UnionWith(names);
                    }
                    if (related.Count == 0)
                        continue;
                    string locking;
                    try
                    {
                        locking = dimension.NumberOfSegments > 0
                            ? "цепочка; EQ=" + dimension.AreSegmentsEqual + "; замки=" + string.Join(",", dimension.Segments.Cast<DimensionSegment>().Select(s => s.IsLocked))
                            : "замок=" + dimension.IsLocked;
                    }
                    catch (Exception ex) { locking = "замки не прочитаны: " + OperationError.Message(ex); }
                    string description = "Размер / привязка [ID " + Common.IDHelper.ElIdValue(dimension.Id) + "] [" + locking + "; " + labelStatus + "] → "
                        + string.Join("; ", references.Select(r => ReferenceDescription(document, r)));
                    foreach (string name in related)
                        if (result.Parameters.ContainsKey(name))
                            result.Parameters[name].Links.Add(description);
                }
                catch (Exception ex) { result.ReadWarnings.Add("Не удалось прочитать размер [ID " + Common.IDHelper.ElIdValue(dimension.Id) + "]: " + OperationError.Message(ex)); }
            }
            return result;
        }

        private static string ReferenceDescription(Document document, Reference reference)
        {
            var element = document.GetElement(reference.ElementId);
            string description = DependencyElement(element);
            try
            {
                var instance = element as FamilyInstance;
                if (instance != null)
                    description += " / опора «" + instance.GetReferenceName(reference) + "»";
                description += " / " + reference.ConvertToStableRepresentation(document);
            }
            catch (Exception ex) { description += " / опора не прочитана: " + OperationError.Message(ex); }
            return description;
        }

        private static string ReadNestedDiagnostics(Document document, InstallationConfiguration configuration)
        {
            var lines = new List<string> { "Вложенные семейства: параметры размеров, формулы и привязки (без изменения файлов)." };
            var requested = configuration.Blocks.Select(b => b.Type).ToList();
            var manager = document.FamilyManager;
            var sourceType = manager.CurrentType ?? manager.Types.Cast<FamilyType>().FirstOrDefault(t => t.Name == BaseFamilyTypeName);
            if (sourceType != null)
                foreach (int slot in new[] { 1, 2, 12 })
                {
                    var original = DescribeType(document, sourceType.AsElementId(SectionParameter(manager, slot)));
                    requested.Add(original);
                    lines.Add("Исходный тип «" + InstallationConfiguration.ParameterName(slot) + "»: " + original.DisplayName);
                }
            string[] sizes = { "КП_Р_Длина", "КП_Р_Ширина", "КП_Р_Высота", "КП_Р_Смещение", "Видимые" };
            var families = new FilteredElementCollector(document).OfClass(typeof(Family)).Cast<Family>().ToList();
            foreach (var group in requested.GroupBy(c => c.FamilyName))
            {
                var family = families.SingleOrDefault(f => f.Name == group.Key);
                if (family == null || !family.IsEditable)
                {
                    lines.Add("Недоступно для чтения: " + group.Key);
                    continue;
                }
                Document nested = null;
                bool closed = true;
                try
                {
                    nested = document.EditFamily(family);
                    foreach (var choice in group.GroupBy(c => c.Key).Select(g => g.First()))
                    {
                        lines.Add("\nВложенный тип: " + choice.DisplayName);
                        var type = nested.FamilyManager.Types.Cast<FamilyType>().FirstOrDefault(t => t.Name == choice.TypeName);
                        if (type == null)
                        {
                            lines.Add("Тип не найден в редактируемом вложенном семействе.");
                            continue;
                        }
                        lines.Add(ReadFamilyDependencies(nested, type, sizes).Report(sizes));
                    }
                }
                catch (Exception ex) { lines.Add("Не удалось прочитать «" + group.Key + "»: " + OperationError.Message(ex)); }
                finally
                {
                    if (nested != null && nested.IsValidObject)
                        try
                        {
                            closed = nested.Close(false);
                        }
                        catch (Exception ex) { closed = false; lines.Add("Закрытие вложенного документа: " + OperationError.Message(ex)); }
                }
                if (!closed)
                {
                    lines.Add("Не удалось закрыть временный вложенный документ «" + group.Key + "».");
                    break;
                }
            }
            return string.Join("\n", lines);
        }

        private static FamilyParameter SectionParameter(FamilyManager manager, int slot)
        {
            string name = InstallationConfiguration.ParameterName(slot);
            var parameter = FindFamilyParameter(manager, name);
            if (parameter == null || parameter.StorageType != StorageType.ElementId)
                throw new InvalidOperationException("В исходном семействе нет параметра состава «" + name + "».");
            return parameter;
        }

        private static SectionTypeChoice DescribeType(Document document, ElementId id)
        {
            var element = document.GetElement(id);
            var type = element as ElementType;
            if (type != null)
                return new SectionTypeChoice(type.FamilyName, type.Name);
            var nested = element as NestedFamilyTypeReference;
            if (nested != null)
                return new SectionTypeChoice(nested.FamilyName, nested.TypeName);
            throw new InvalidOperationException("Не удалось прочитать название вложенного типа оборудования.");
        }

        private static List<KeyValuePair<ElementId, SectionTypeChoice>> ReadTypeChoices(Document document, FamilyParameter parameter)
        {
            return document.OwnerFamily.GetFamilyTypeParameterValues(parameter.Id)
                .Select(id => new KeyValuePair<ElementId, SectionTypeChoice>(id, DescribeType(document, id))).ToList();
        }

        private static void ValidateDiskSource(UIApplication app, string path)
        {
            foreach (Document open in app.Application.Documents)
                if (open.IsValidObject && SamePath(open.PathName, path) && open.IsModified)
                    throw new InvalidOperationException("В Revit есть несохранённые изменения этого семейства. Сохраните их перед чтением файла в конфигураторе.");
        }

        private static FamilyPackage ReadFamilyPackage(UIApplication app, string path, bool includeTypes, bool allowOpenChanges = false)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("Файл семейства недоступен.", path);
            if (!allowOpenChanges)
                ValidateDiskSource(app, path);
            string directory = Path.Combine(Path.GetTempPath(), "KPLN_Ventilation_Read_" + Guid.NewGuid().ToString("N"));
            Document document = null;
            try
            {
                Directory.CreateDirectory(directory);
                string input = Path.Combine(directory, Path.GetFileName(path));
                File.Copy(path, input, false);
                File.SetAttributes(input, File.GetAttributes(input) & ~FileAttributes.ReadOnly);
                document = app.Application.OpenDocumentFile(input);
                if (!document.IsFamilyDocument)
                    throw new InvalidOperationException("Выбранный файл не является семейством Revit.");
                bool library = SamePath(path, ResolveSourcePath()) || SamePath(path, SourceFamilyPath) || SamePath(path, LiteralSourceFamilyPath);
                var result = ReadFamilyPackage(document, library ? null : path, includeTypes && !library);
                if (!document.Close(false))
                    throw new InvalidOperationException("Не удалось закрыть временный документ после чтения типов.");
                document = null;
                return result;
            }
            finally
            {
                bool closed = document == null || !document.IsValidObject;
                if (!closed)
                {
                    try
                    {
                        closed = document.Close(false);
                    }
                    catch { closed = false; }
                }
                if (closed && Directory.Exists(directory))
                {
                    try
                    {
                        Directory.Delete(directory, true);
                    }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
        }

        private static FamilyPackage ReadOpenFamilyPackage(UIApplication app, Document document, string path)
        {
            var family = ReadFamilyPackage(document, path, true);
            if (!document.IsModified)
                return family;
            var saved = ReadFamilyPackage(app, path, true, true);
            foreach (var type in family.Types)
            {
                var disk = saved.Types.FirstOrDefault(t => t.PersistedName == type.PersistedName);
                if (type.Configuration == null || disk?.Configuration == null
                    || type.Configuration.Signature() != disk.Configuration.Signature())
                    type.RequireSave();
            }
            return family;
        }

        private static FamilyPackage ReadFamilyPackage(Document document, string path, bool includeTypes)
        {
            var manager = document.FamilyManager;
            var source = manager.Types.Cast<FamilyType>().SingleOrDefault(t => t.Name == BaseFamilyTypeName);
            if (source == null)
                throw new InvalidOperationException("В семействе нет исходного типа «" + BaseFamilyTypeName
                + "». Нельзя прочитать начальные значения для его копии.");
            var result = new FamilyPackage
            {
                Path = path,
                SourcePath = path,
                SelectedTypeName = manager.CurrentType?.Name,
                Catalog = ReadSectionCatalog(document, source)
            };
            // Ошибка необязательной проверки основы не должна мешать чтению типов.
            try
            {
                result.DefinitionSnapshot = ReadFamilyDefinitionSnapshot(document, result.Catalog);
            }
            catch { result.DefinitionSnapshot = null; }
            if (includeTypes)
                foreach (var type in manager.Types.Cast<FamilyType>().Where(t => t.Name != BaseFamilyTypeName).OrderBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase))
                {
                    var item = new FamilyTypeItem { Name = type.Name, PersistedName = type.Name, SavedPath = path, SourcePath = path };
                    try
                    {
                        item.Configuration = ReadConfiguration(document, type, result.Catalog);
                        item.HasAutomaticName = type.Name == item.Configuration.Info.AutomaticTypeName;
                        item.MarkSaved();
                    }
                    catch (Exception ex) { item.LoadError = "Не удалось прочитать тип «" + type.Name + "»: " + OperationError.Message(ex); }
                    result.Types.Add(item);
                }
            return result;
        }

        private static SectionCatalog ReadSectionCatalog(Document document, FamilyType sourceType, SectionCatalog choicesCatalog = null)
        {
            var manager = document.FamilyManager;
            var countParameter = manager.get_Parameter("Секции_Промежуточные_Количество");
            var valveParameter = manager.get_Parameter("Соединитель_Приточный_Клапан");
            if (countParameter == null || valveParameter == null || countParameter.StorageType != StorageType.Integer
                || valveParameter.StorageType != StorageType.Integer)
                throw new InvalidOperationException("Не найдены параметры количества секций и наличия клапана.");
            int? count = sourceType.AsInteger(countParameter), valve = sourceType.AsInteger(valveParameter);
            if (!count.HasValue || count < 1 || count > 9 || !valve.HasValue || (valve != 0 && valve != 1))
                throw new InvalidOperationException("В типе «" + sourceType.Name + "» некорректное количество секций или состояние клапана.");
            var catalog = new SectionCatalog
            {
                SourceTypeName = sourceType.Name,
                InstallationWidthMm = ReadSourceLengthMm(manager, sourceType, "Установка_Ширина"),
                InstallationHeightMm = ReadSourceLengthMm(manager, sourceType, "Установка_Высота"),
                FrameHeightMm = ReadSourceLengthMm(manager, sourceType, "Основание_Рама_Высота"),
                IntermediateCount = count.Value,
                HasValve = valve.Value != 0
            };
            for (int slot = 1; slot <= 12; slot++)
            {
                var parameter = SectionParameter(document.FamilyManager, slot);
                var definition = new SectionDefinition { ParameterName = parameter.Definition.Name };
                if (choicesCatalog == null)
                {
                    var values = ReadTypeChoices(document, parameter);
                    definition.Choices = SectionDefinition.OrderChoices(values.Select(v => v.Value));
                    if (definition.Choices.Count == 0)
                        throw new InvalidOperationException("Нет доступных типов для «" + parameter.Definition.Name + "».");
                    string category = document.GetElement(values[0].Key).Category?.Name;
                    definition.DisplayName = parameter.Definition.Name + (string.IsNullOrWhiteSpace(category) ? "" : "<" + category + ">");
                }
                else
                {
                    definition.Choices = choicesCatalog[slot].Choices;
                    definition.DisplayName = choicesCatalog[slot].DisplayName;
                }
                var selectedId = sourceType.AsElementId(parameter);
                if (selectedId != null && !selectedId.Equals(ElementId.InvalidElementId))
                {
                    var selected = DescribeType(document, selectedId);
                    definition.SourceType = definition.Choices.SingleOrDefault(c => c.Key == selected.Key);
                    if (definition.SourceType == null)
                        throw new InvalidOperationException("Исходный тип «" + selected.DisplayName
                        + "» недоступен для параметра «" + parameter.Definition.Name + "».");
                }
                string lengthName = InstallationConfiguration.DimensionParameterName(slot, "Длина");
                definition.DefaultLengthMm = ReadSourceLengthMm(manager, sourceType, lengthName);
                if (slot == 1 || slot == 12)
                {
                    definition.DefaultWidthMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Ширина"));
                    definition.DefaultHeightMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Высота"));
                    definition.DefaultOffsetXMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Смещение по X"));
                    definition.DefaultOffsetYMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Смещение по Y"));
                }
                catalog.Add(slot, definition);
            }
            foreach (string name in SharedParameters.LengthNames)
                catalog.SharedLengthsMm.Add(name, ReadSourceLengthMm(manager, sourceType, name));
            foreach (string name in SharedParameters.BooleanNames)
                catalog.SharedFlags.Add(name, ReadBooleanValue(manager, sourceType, name));
            catalog.Dependencies = ReadCatalogDependencies(document, sourceType, choicesCatalog?.Dependencies);
            return catalog;
        }

        private static InstallationConfiguration ReadConfiguration(Document document, FamilyType type, SectionCatalog catalog)
        {
            // Та же процедура чтения, но снимок значений принадлежит именно выбранному типу.
            var configuration = InstallationConfiguration.CreateDraft(ReadSectionCatalog(document, type, catalog));
            var manager = document.FamilyManager;
            configuration.Info = new InstallationInfo
            {
                SystemName = ReadInfoValue(manager, type, InstallationInfo.SystemNameParameter),
                Manufacturer = ReadInfoValue(manager, type, InstallationInfo.ManufacturerParameter),
                Mark = ReadInfoValue(manager, type, InstallationInfo.MarkParameter),
                Unit = ReadInfoValue(manager, type, InstallationInfo.UnitParameter),
                Description = ReadInfoValue(manager, type, InstallationInfo.DescriptionParameter),
                ProductCode = ReadInfoValue(manager, type, InstallationInfo.ProductCodeParameter),
                MassText = ReadInfoValue(manager, type, InstallationInfo.MassTextParameter)
            };
            return configuration;
        }

        private static void ApplySectionTypes(Document document, InstallationConfiguration configuration)
        {
            FamilyManager manager = document.FamilyManager;
            foreach (var assignment in configuration.SectionAssignments())
            {
                FamilyParameter parameter = SectionParameter(manager, assignment.Key);
                // Идентификаторы разрешаем в создаваемой копии, а не переносим из закрытого документа каталога.
                var matches = ReadTypeChoices(document, parameter).Where(v => v.Value.Key == assignment.Value.Key).ToList();
                if (matches.Count != 1)
                    throw new InvalidOperationException("Не найден однозначный тип «" + assignment.Value.DisplayName + "» для «" + parameter.Definition.Name + "».");
                var current = manager.CurrentType.AsElementId(parameter);
                if (current != null && current.Equals(matches[0].Key))
                    continue;
                if (parameter.IsDeterminedByFormula)
                    continue;
                EnsureWritable(parameter);
                try
                {
                    manager.Set(parameter, matches[0].Key);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "», тип «"
                    + assignment.Value.DisplayName + "»: " + OperationError.Message(ex), ex);
                }
            }
            // Размеры записываются отдельно. Текстовые параметры и механизмы здесь не меняем.
        }
    }
}
