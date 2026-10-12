using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using KPLN_Tools_OVVK.Forms.ViewModels;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal sealed partial class VentFamilyService
    {
        internal bool IsClosed
        { get; private set; }

        internal VentFamilyService(OV_VentConfiguratorVM owner)
        {
            _owner = owner;
        }

        internal void Close()
        {
            IsClosed = true;
        }

        private static readonly Dictionary<string, string> SavedFamilyPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _automaticReadDirectories = new List<string>();
        private string _activeContextKey;
        private readonly Dictionary<string, PendingProjectImport> _projectImports = new Dictionary<string, PendingProjectImport>(StringComparer.Ordinal);

        private static string ContextKey(Document document, ProjectItem project)
        {
            if (project != null)
                return "project:" + project.Name + "|" + project.Stage + "|" + project.Code;
            if (document == null || !document.IsValidObject)
                return "no-document";
            string path = GetModelPaths(document).FirstOrDefault(p => !string.IsNullOrWhiteSpace(p));
            if (string.IsNullOrWhiteSpace(path))
                path = document.PathName;
            return !string.IsNullOrWhiteSpace(path) ? "document:" + path : "unsaved:" + document.GetHashCode();
        }

        private bool UpdateAutomaticFamily(UIApplication app, Document document, ProjectItem project, bool force = false)
        {
            string key = ContextKey(document, project);
            _activeContextKey = key;
            if (_owner.SwitchFamilyContext(key) && !force)
                return true;
            _owner.SetBusy(true);
            try
            {
                FamilyPackage package = null;
                string path = ProjectFamilyStorage.GetPath(project);
                if (path != null && !EnsureProjectFamily(path))
                    return false;
                string remembered;
                if (path == null && SavedFamilyPaths.TryGetValue(key, out remembered) && File.Exists(remembered))
                    path = remembered;
                if (path != null)
                {
                    var loaded = FindLoadedInstallationFamily(document, project?.Code);
                    // ImportProjectFamily уже читает RFA: не открываем его второй раз заранее.
                    if (loaded != null)
                        package = ImportProjectFamily(app, document, loaded, path);
                    else
                    {
                        var opened = FindOpenFamily(app, path);
                        package = opened == null ? ReadFamilyPackage(app, path, true) : ReadOpenFamilyPackage(app, opened, path);
                    }
                }
                else if (document != null && document.IsValidObject && document.IsFamilyDocument
                    && document.FamilyManager.get_Parameter("Секции_Промежуточные_Количество") != null)
                {
                    if (string.IsNullOrWhiteSpace(document.PathName) || !File.Exists(document.PathName))
                        throw new InvalidOperationException("Сначала сохраните открытое семейство в Revit. После этого его типы можно прочитать автоматически.");
                    bool library = SamePath(document.PathName, ResolveSourcePath());
                    package = library ? ReadFamilyPackage(document, null, false) : ReadOpenFamilyPackage(app, document, document.PathName);
                }
                else if (document != null && document.IsValidObject && !document.IsFamilyDocument)
                {
                    var family = FindLoadedInstallationFamily(document, project?.Code);
                    if (family != null)
                        package = ReadProjectFamily(app, document, family);
                }
                if (package == null)
                    package = ReadFamilyPackage(app, ResolveSourcePath(), false);
                package.UpdateNotice = CheckProjectFamilyUpdates(app, package.Path, package.DefinitionSnapshot);
                _owner.SetFamily(package);
                if (package.LoadNotice != null)
                    _owner.SetStatus(package.LoadNotice, false);
                return true;
            }
            catch (System.OperationCanceledException)
            {
                _owner.SetStatus("Чтение отменено. Повторите запуск конфигуратора. Файлы не изменены.", false);
                return false;
            }
            catch (Exception ex)
            {
                _owner.SetOperationError(OperationError.Create("Автоматическое чтение типов семейства", ex, null, false));
                return false;
            }
            finally { _owner.SetBusy(false); }
        }

        private readonly OV_VentConfiguratorVM _owner;
        private UIApplication _application;
        private IList<ProjectItem> _projects = new List<ProjectItem>();
        private string _lastModelKey;

        internal void TryUpdateActiveProject(UIApplication app, bool force)
        {
            string filePath = string.Empty;
            try
            {
                Document document = app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document;
                // Читаем фактический путь отдельно от сопоставления проекта, в том числе для RFA.
                if (document != null && document.IsValidObject)
                    filePath = document.PathName ?? string.Empty;
                UpdateActiveProject(document, filePath, force);
            }
            catch (Exception ex)
            {
                // Ошибка чтения пути не должна оставлять в заголовке предыдущий проект.
                string key = "Ошибка определения проекта: " + ex.Message;
                if (string.Equals(_lastModelKey, key, StringComparison.Ordinal))
                    return;
                _lastModelKey = key;
                _owner.SetProject(new ProjectMatch());
                _owner.SetStatus(key, true);
            }
        }

        internal void RefreshProjects(UIApplication app)
        {
            _application = app;
            string databaseError = null;
            try
            {
                _projects = ProjectDatabase.Read(ProjectDatabasePath);
            }
            catch (Exception ex)
            {
                _projects = new List<ProjectItem>();
                databaseError = ex.Message;
            }

            TryUpdateActiveProject(app, true);
            if (databaseError != null && !_owner.HasOperationError)
                _owner.SetStatus(databaseError, true);
        }

        private static string[] GetModelPaths(Document document)
        {
            if (document == null || !document.IsValidObject || document.IsFamilyDocument || document.IsLinked)
                return new[] { string.Empty, string.Empty };
            string centralPath = string.Empty;
            if (document.IsWorkshared)
            {
                try
                {
                    using (ModelPath modelPath = document.GetWorksharingCentralModelPath())
                        if (modelPath != null)
                            centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException) { }
            }
            return new[] { centralPath, document.PathName ?? string.Empty };
        }

        private ProjectMatch MatchModel(string[] paths)
        {
            // Центральный путь приоритетен. К локальному переходим только при отсутствии совпадений.
            var central = ProjectPathMatcher.Find(_projects, paths[0]);
            return central.Project != null || central.IsAmbiguous
                ? central : ProjectPathMatcher.Find(_projects, paths[1]);
        }

        private ProjectMatch MatchDocument(Document document)
        {
            // Открытое проектное RFA сохраняет привязку по своей папке и аббревиатуре.
            if (document != null && document.IsValidObject && document.IsFamilyDocument)
                return ProjectFamilyStorage.Find(_projects, document.PathName);
            return MatchModel(GetModelPaths(document));
        }

        private void UpdateActiveProject(Document document, string filePath, bool force)
        {
            string[] paths = GetModelPaths(document);
            string displayPath = !string.IsNullOrWhiteSpace(filePath) ? filePath
                : !string.IsNullOrWhiteSpace(paths[0]) ? paths[0]
                : document == null || !document.IsValidObject ? "Нет открытого файла" : "Файл не сохранён";
            string key = displayPath + "\n" + string.Join("\n", paths) + "\n" + (document == null ? "" : document.GetHashCode().ToString());
            if (!force && string.Equals(_lastModelKey, key, StringComparison.Ordinal))
                return;
            _lastModelKey = key;
            var match = MatchDocument(document);
            _owner.SetProject(match);
            UpdateAutomaticFamily(_application, document, match.Project);
        }

        private static IList<OpenDocumentItem> GetOpenProjects(UIApplication app)
        {
            var items = new List<OpenDocumentItem>();
            foreach (Document doc in app.Application.Documents)
            {
                if (!doc.IsValidObject || doc.IsFamilyDocument || doc.IsLinked || doc.IsReadOnly)
                    continue;
                items.Add(new OpenDocumentItem
                {
                    Document = doc,
                    DisplayName = doc.Title,
                    LoadTargetName = doc.Title + (string.IsNullOrWhiteSpace(doc.PathName)
                        ? " — не сохранён" : " — " + doc.PathName)
                });
            }
            return items.OrderBy(p => p.DisplayName).ToList();
        }
    }
}
