using System;
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
        internal void Execute(UIApplication app, FamilyRequest request)
        {
            if (IsClosed)
                return;
            string savedPath = null;
            string stage = "Проверка настроек";
            string diagnostics = null;
            try
            {
                if (request == null)
                    return;
                if (request.Kind == RequestKind.LoadSectionCatalog)
                {
                    stage = "Чтение состава исходного семейства";
                    Document active = app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document;
                    if (UpdateAutomaticFamily(app, active, MatchDocument(active).Project, true))
                        _owner.CreateTypeIfReady();
                    return;
                }

                Document activeDocument = app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document;
                ProjectItem project = MatchDocument(activeDocument).Project;
                // Между нажатием кнопки и обработкой очереди пользователь мог сменить документ.
                if (!string.Equals(request.ContextKey, ContextKey(activeDocument, project), StringComparison.OrdinalIgnoreCase))
                {
                    TryUpdateActiveProject(app, true);
                    throw new InvalidOperationException("Активный проект изменился. Выберите нужный тип в текущем проекте и повторите действие.");
                }

                if (request.Kind == RequestKind.UpdateFamily)
                {
                    stage = "Обновление семейства проекта из основы";
                    if (!CanUpdateProjectFamily())
                        throw new UnauthorizedAccessException("Обновление семейства доступно только отделу 8.");
                    string path = ProjectFamilyStorage.GetPath(project) ?? request.OutputPath;
                    string backup;
                    var updated = UpdateEntireProjectFamily(app, path, request.TypeName, request.Types, value => stage = value, out backup);
                    if (updated == null)
                    {
                        _owner.SetStatus("Обновление отменено. Семейство проекта не изменено.", false);
                        return;
                    }

                    savedPath = path;
                    if (_activeContextKey != null)
                    {
                        SavedFamilyPaths[_activeContextKey] = path;
                        _projectImports.Remove(_activeContextKey);
                    }

                    _owner.SetFamily(updated);
                    _owner.SetStatus("Семейство проекта обновлено из основы. Существующие типы сохранены. Резервная копия: " + backup, false);
                    return;
                }

                bool deleting = request.Kind == RequestKind.DeleteType;
                string typeName = deleting ? request.PersistedName : VentTypeName.ValidateTypeName(request.TypeName);
                if (!deleting)
                {
                    if (request.Configuration == null)
                        throw new InvalidOperationException("Сначала создайте тип.");
                    request.Configuration.ValidateForOutput();
                }

                string configuredPath = ProjectFamilyStorage.GetPath(project);
                var action = FamilyActionPlan.Make(request.Kind, request.OutputPath, configuredPath, request.PersistedName, request.IsDirty);
                // Открытие и загрузка не заходят в сохранение и не меняют RFA на диске.
                if (!action.WritesFamily)
                {
                    string path = action.Path;
                    if (!File.Exists(path))
                        throw new FileNotFoundException("Сохранённое семейство недоступно. Проверьте путь к файлу.", path);
                    if (request.Kind == RequestKind.OpenInRevit)
                    {
                        stage = "Открытие сохранённой установки в Revit";
                        OpenSavedFamily(app, path, request.PersistedName, value => diagnostics = value);
                        _owner.SetStatus("Вент. установка открыта с типом «" + request.PersistedName + "».", false);
                    }
                    else
                    {
                        stage = "Выбор проекта для загрузки";
                        var projects = GetOpenProjects(app);
                        if (projects.Count == 0)
                            throw new InvalidOperationException("Нет доступных проектов. Откройте проект в Revit и повторите загрузку.");
                        var chosen = _owner.ChooseLoadProject(projects, activeDocument);
                        if (chosen == null)
                        {
                            _owner.SetStatus("Выбор проекта отменён.", false);
                            return;
                        }

                        var target = GetTargetDocument(app, chosen);
                        stage = "Загрузка сохранённого типа в проект";
                        _owner.SetStatus(LoadSavedType(app, target, path, request.PersistedName), false);
                    }

                    return;
                }

                string sourcePath = ResolveSourcePath();
                stage = "Определение файла для сохранения";
                string outputPath = action.Path;
                if (action.NeedsPath)
                    outputPath = _owner.ChooseOutputPath(project?.Code, null);
                if (outputPath == null)
                {
                    _owner.SetStatus("Сохранение отменено. Файлы не изменены.", false);
                    return;
                }

                if (configuredPath != null && !EnsureProjectFamily(configuredPath))
                    return;
                stage = "Проверка пути сохранения";
                outputPath = ValidateOutputPath(app, sourcePath, outputPath, configuredPath);
                var openFamily = FindOpenFamily(app, outputPath);
                PendingProjectImport import;
                _projectImports.TryGetValue(_activeContextKey, out import);
                if (import != null)
                {
                    import.CheckDestination(outputPath);
                    if (openFamily != null)
                        throw new InvalidOperationException("Закройте выбранный RFA в Revit перед записью перенесённых типов.");
                }

                string inputPath = import != null ? import.SnapshotPath : File.Exists(outputPath) ? outputPath : !string.IsNullOrWhiteSpace(request.FamilyPath) ? request.FamilyPath : sourcePath;
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException("Файл семейства недоступен. Проверьте путь и подключение к сети.", inputPath);
                if (openFamily == null)
                    ValidateDiskSource(app, inputPath);
                bool freshLibraryCopy = import?.FromLibrary == true || SamePath(inputPath, sourcePath) || SamePath(inputPath, SourceFamilyPath) || SamePath(inputPath, LiteralSourceFamilyPath);
                string persistedName = freshLibraryCopy && !deleting ? null : request.PersistedName;
                FamilyPackage existing = null;
                if (!freshLibraryCopy && File.Exists(outputPath))
                {
                    stage = "Чтение существующих типов";
                    existing = openFamily == null ? ReadFamilyPackage(app, inputPath, true) : ReadFamilyPackage(openFamily, outputPath, true);
                    if (!deleting)
                    {
                        if (persistedName == null && request.HasAutomaticName)
                            typeName = _owner.AvailableAutomaticName(typeName, existing.Types.Select(t => t.Name));
                        if (import == null)
                            _owner.MergeDiscoveredFamily(existing, typeName);
                        if (persistedName == null && existing.Types.Any(t => string.Equals(t.Name, typeName, StringComparison.OrdinalIgnoreCase)))
                            throw new InvalidOperationException("Тип «" + typeName + "» уже есть в файле и добавлен в список. Выберите его для редактирования или задайте другое имя новому типу.");
                    }
                }

                FamilyPackage savedFamily;
                string warning;
                bool unchanged = import == null && !deleting && !request.IsDirty && persistedName != null && SamePath(outputPath, request.OutputPath) && existing != null;
                if (unchanged && existing != null && !existing.Types.Any(t => t.PersistedName == persistedName))
                    throw new InvalidOperationException("В семействе больше нет типа «" + persistedName + "».");
                // Явное «Сохранить» исправляет и ранее созданный тип с длинами скрытых блоков из шаблона.
                if (unchanged && existing.Types.Any(t => t.PersistedName == persistedName && (t.Configuration == null || t.Configuration.NeedsInactiveLengthReset)))
                    unchanged = false;
                if (unchanged && openFamily == null)
                {
                    savedFamily = existing;
                    warning = string.Empty;
                }
                else if (openFamily != null)
                {
                    savedFamily = SaveOpenFamily(openFamily, outputPath, typeName, persistedName, request.Configuration, deleting, value => stage = value, value => diagnostics = value, !unchanged);
                    warning = string.Empty;
                }
                else
                    warning = BuildAndSaveFamily(app, inputPath, outputPath, typeName, persistedName, freshLibraryCopy, request.Configuration, value => stage = value, value => diagnostics = value, out savedFamily, deleting, import);
                savedPath = outputPath;
                if (import != null)
                    _projectImports.Remove(_activeContextKey);
                if (_activeContextKey != null)
                    SavedFamilyPaths[_activeContextKey] = outputPath;
                if (deleting)
                {
                    _owner.CompleteDeletedFamily(savedFamily, request.TargetType);
                    _owner.SetStatus("Тип «" + persistedName + "» удалён из семейства. Файл сохранён." + warning, false);
                }
                else
                {
                    _owner.ApplyCalculatedDimensions(request.Configuration);
                    _owner.RememberSavedFamily(savedFamily, typeName);
                    _owner.SetStatus("Вент. установка сохранена с типом «" + typeName + "». " + savedPath + warning, false);
                }
            }
            catch (System.OperationCanceledException)
            {
                _owner.SetStatus("Выбор семейства отменён. Файлы не изменены.", false);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                _owner.SetStatus("Операция отменена.", false);
            }
            catch (Exception ex)
            {
                var error = OperationError.Create(stage, ex, savedPath, request != null && (request.Kind == RequestKind.Save || request.Kind == RequestKind.DeleteType || request.Kind == RequestKind.UpdateFamily));
                _owner.SetOperationError(error);
            }
            finally
            {
                _owner.SetBusy(false);
            }
        }
    }
}
