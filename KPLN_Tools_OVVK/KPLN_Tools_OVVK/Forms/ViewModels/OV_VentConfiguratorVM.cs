using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using KPLN_Tools_OVVK.Forms.Entities;
using KPLN_Tools_OVVK.Forms.Models.Core;
using KPLN_Tools_OVVK.Services.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Forms.ViewModels
{
    internal sealed partial class OV_VentConfiguratorVM
    {
        private readonly OV_VentConfiguratorForm _mainWindow;

        internal VentFamilyService FamilyService
        { get; }

        public OV_VentConfiguratorM CurrentOV_VentConfiguratorM { get; } = new OV_VentConfiguratorM();

        private OV_VentConfiguratorM Model
        {
            get
            {
                return CurrentOV_VentConfiguratorM;
            }
        }

        public ICommand CreateTypeCmd
        { get; }

        public ICommand CopyTypeCmd
        { get; }

        public ICommand UpdateFamilyCmd
        { get; }

        public ICommand DeleteTypeCmd
        { get; }

        public ICommand SaveCmd
        { get; }

        public ICommand OpenInRevitCmd
        { get; }

        public ICommand AddToProjectCmd
        { get; }

        public ICommand CloseWindowCmd
        { get; }

        internal OV_VentConfiguratorVM(OV_VentConfiguratorForm mainWindow)
        {
            _mainWindow = mainWindow;
            FamilyService = new VentFamilyService(this);
            CreateTypeCmd = new RelayCommand<object>(_ => CreateTypeOrLoadCatalog(), _ => Model.CanEdit);
            CopyTypeCmd = new RelayCommand<object>(_ => CopyType(), _ => Model.HasConfiguration);
            UpdateFamilyCmd = new RelayCommand<object>(_ => UpdateFamily(), _ => Model.CanUpdateFamily);
            DeleteTypeCmd = new RelayCommand<FamilyTypeItem>(DeleteType, item => Model.CanEdit && item != null && !item.IsCreate);
            SaveCmd = new RelayCommand<object>(_ => QueueRequest(RequestKind.Save), _ => Model.HasConfiguration);
            OpenInRevitCmd = new RelayCommand<object>(_ => QueueRequest(RequestKind.OpenInRevit), _ => Model.HasConfiguration);
            AddToProjectCmd = new RelayCommand<object>(_ => QueueRequest(RequestKind.AddToProject), _ => Model.HasConfiguration);
            CloseWindowCmd = new RelayCommand<object>(_ => _mainWindow.Close(), _ => Model.CanEdit);
        }

        internal void RenameType(string name)
        {
            if (Model.IsSynchronizing || Model.IsBusy || Model.CurrentType == null || Model.CurrentType.Name == name)
                return;
            Model.CurrentType.Name = name;
            Model.CurrentType.MarkChanged();
            Model.CurrentType.HasAutomaticName = false;
        }

        internal void SelectSection(SectionValue section)
        {
            Model.SelectedSection = section;
            _mainWindow.RenderBlocks();
            _mainWindow.RenderSelectedSection();
        }

        internal void ResizeSections(int count)
        {
            if (Model.IsSynchronizing || Model.IsBusy || Model.Configuration == null || count == Model.Configuration.SectionCount)
                return;
            EditSections(() => Model.Configuration.Resize(count));
        }

        internal void SetValve(bool enabled)
        {
            if (Model.IsSynchronizing || Model.IsBusy || Model.Configuration == null)
                return;
            EditSections(() =>
            {
                Model.Configuration.SetValve(enabled);
                if (enabled)
                    Model.SelectedSection = Model.Configuration.Blocks[1];
            });
        }

        internal void Dispose()
        {
            FamilyService.Close();
            foreach (var section in _observedSections)
                section.PropertyChanged -= SectionChanged;
            _observedSections.Clear();
            ClearDimensionObservers();
            ClearInfoObserver();
            Model.Workspaces.Clear();
        }

        internal string ChooseExistingFamilyPath(string familyName, string failure)
        {
            return _mainWindow.ChooseExistingFamilyPath(familyName, failure);
        }

        internal string ChooseOutputPath(string projectCode, string familyPath)
        {
            return _mainWindow.ChooseOutputPath(projectCode, familyPath);
        }

        internal OpenDocumentItem ChooseLoadProject(IList<OpenDocumentItem> projects, Autodesk.Revit.DB.Document preferred)
        {
            return _mainWindow.ChooseLoadProject(projects, preferred);
        }

        internal bool IsBusy
        {
            get
            {
                return Model.IsBusy;
            }
        }

        internal bool HasOperationError
        {
            get
            {
                return Model.HasOperationError;
            }
        }

        internal void CreateTypeIfReady()
        {
            if (Model.SectionCatalog != null)
                CreateType();
        }

        internal void SetProject(ProjectMatch match)
        {
            // В заголовке показываем рабочее семейство, а не путь активной модели RVT.
            Model.HeaderProject = match.Project;
            UpdateProjectHeader();
        }

        private void UpdateProjectHeader()
        {
            bool known = Model.HeaderProject != null && !Model.HeaderProject.IsUnknown;
            string project = known ? Model.HeaderProject.DisplayName : UnknownProjectName;
            string familyPath = known ? null : Model.LastSavedPath ?? Model.WorkingFamilyPath;
            if (known)
            {
                try
                {
                    familyPath = ProjectFamilyStorage.GetPath(Model.HeaderProject);
                }
                // Ошибку аббревиатуры отдельно выводит загрузчик семейства.
                catch (InvalidOperationException)
                {
                }
            }

            Model.ProjectHeader = project + (string.IsNullOrWhiteSpace(familyPath) ? string.Empty : " - " + familyPath);
            Model.RefreshState();
        }

        private void CreateTypeOrLoadCatalog()
        {
            if (Model.IsBusy)
                return;
            if (Model.SectionCatalog == null)
                QueueRequest(RequestKind.LoadSectionCatalog);
            else
                CreateType();
        }

        private void CopyType()
        {
            if (Model.IsBusy || Model.CurrentType?.Configuration == null)
                return;
            var copy = Model.CurrentType.CreateCopy(Model.Types);
            copy.SelectedTab = _mainWindow.SelectedTab;
            copy.SelectedSectionIndex = Math.Max(0, Model.Configuration.Blocks.IndexOf(Model.SelectedSection));
            Model.Types.Add(copy);
            Model.IsSynchronizing = true;
            try
            {
                _mainWindow.SelectedType = copy;
            }
            finally
            {
                Model.IsSynchronizing = false;
            }

            SelectType(copy);
            SetStatus("Создана копия типа. Укажите имя и нажмите «Сохранить».", false);
        }

        internal bool HasUnsavedEdits
        {
            get
            {
                return Model.Types.Any(t => !t.IsCreate && t.Configuration != null && t.HasUserChanges);
            }
        }

        internal bool HasUnsavedTypes
        {
            get
            {
                foreach (var item in Model.Types.Where(t => !t.IsCreate && t.Configuration != null))
                    item.MarkChanged();
                return Model.Types.Any(t => !t.IsCreate && t.Configuration != null && t.IsDirty);
            }
        }

        private void CreateType()
        {
            var item = FamilyTypeItem.CreateDraft(Model.Types, Model.SectionCatalog);
            item.SavedPath = Model.WorkingFamilyPath;
            item.SourcePath = Model.WorkingSourcePath;
            Model.Types.Add(item);
            Model.IsSynchronizing = true;
            try
            {
                _mainWindow.SelectedType = item;
            }
            finally
            {
                Model.IsSynchronizing = false;
            }

            SelectType(item);
        }

        internal void SelectType(FamilyTypeItem selected)
        {
            if (selected.Configuration == null)
            {
                Model.IsSynchronizing = true;
                try
                {
                    _mainWindow.SelectedType = Model.CurrentType;
                }
                finally
                {
                    Model.IsSynchronizing = false;
                }

                SetStatus(selected.LoadError ?? "Не удалось прочитать настройки типа.", true);
                return;
            }

            if (Model.CurrentType != null)
            {
                Model.CurrentType.SelectedSectionIndex = Math.Max(0, Model.Configuration.Blocks.IndexOf(Model.SelectedSection));
                Model.CurrentType.SelectedTab = _mainWindow.SelectedTab;
            }

            Model.CurrentType = selected;
            Model.LastSavedPath = selected.SavedPath;
            UpdateProjectHeader();
            Model.IsSynchronizing = true;
            try
            {
                SetConfiguration(selected.Configuration);
                _mainWindow.TypeName = selected.Name;
                _mainWindow.SelectedTab = Math.Max(0, selected.SelectedTab);
                _mainWindow.ShowConfiguration(true);
                _mainWindow.RenderBlocks();
                _mainWindow.RenderSelectedSection();
            }
            finally
            {
                Model.IsSynchronizing = false;
            }

            selected.MarkChanged();
            SetBusy(Model.IsBusy);
            SetStatus("Выбран тип «" + selected.DisplayName + "».", false);
            RecalculateLocally();
        }

        private void RemoveTypeFromList(FamilyTypeItem item)
        {
            int index = Model.Types.IndexOf(item);
            bool active = ReferenceEquals(item, Model.CurrentType);
            Model.IsSynchronizing = true;
            try
            {
                Model.Types.Remove(item);
                if (active)
                {
                    foreach (var block in _observedSections)
                        block.PropertyChanged -= SectionChanged;
                    _observedSections.Clear();
                    ClearDimensionObservers();
                    ClearInfoObserver();
                    Model.CurrentType = null;
                    Model.Configuration = null;
                    Model.SelectedSection = null;
                    Model.LastSavedPath = null;
                    var next = Model.Types.Skip(Math.Min(index, Model.Types.Count)).FirstOrDefault(t => !t.IsCreate && t.Configuration != null) ?? Model.Types.LastOrDefault(t => !t.IsCreate && t.Configuration != null);
                    _mainWindow.SelectedType = next;
                    if (next != null)
                        SelectType(next);
                    else
                    {
                        _mainWindow.ClearConfiguration();
                    }
                }
                else
                    _mainWindow.SelectedType = Model.CurrentType;
            }
            finally
            {
                Model.IsSynchronizing = false;
            }

            SetBusy(false);
            SetStatus("Тип «" + item.DisplayName + "» удалён из списка.", false);
        }

        internal void SetBusy(bool value)
        {
            Model.CanUpdate = CanUpdateProjectFamily();
            Model.IsBusy = value;
            CommandManager.InvalidateRequerySuggested();
        }

        internal void SetStatus(string text, bool error)
        {
            string message = string.IsNullOrWhiteSpace(text) ? (error ? "Ошибка операции." : string.Empty) : text;
            Model.StatusText = error ? "Ошибка" : string.Join(" ", message.Split((char[])null, StringSplitOptions.RemoveEmptyEntries));
            Model.StatusTooltip = message;
            Model.ErrorDetails = error ? message : null;
            Model.HasOperationError = error;
            Model.RefreshState();
        }

        internal void SetOperationError(OperationError error)
        {
            SetStatus(error.Summary, true);
            Model.ErrorDetails = error.Details;
            Model.StatusTooltip = error.Summary;
            Model.RefreshState();
        }

        internal void RememberSavedFamily(FamilyPackage family, string typeName)
        {
            string previous = Model.CurrentType.SavedPath;
            Model.LastSavedPath = Model.WorkingFamilyPath = family.Path;
            Model.WorkingSourcePath = family.SourcePath ?? family.Path;
            UpdateProjectHeader();
            Model.SectionCatalog = family.Catalog;
            Model.CurrentType.Name = typeName;
            Model.CurrentType.PersistedName = typeName;
            Model.CurrentType.SavedPath = family.Path;
            Model.CurrentType.SourcePath = family.SourcePath ?? family.Path;
            Model.CurrentType.MarkSaved();
            foreach (var item in Model.Types.Where(t => !t.IsCreate && t != Model.CurrentType))
                if (item.SavedPath == null || item.SavedPath == previous)
                {
                    item.SavedPath = family.Path;
                    item.SourcePath = family.SourcePath ?? family.Path;
                }

            foreach (var saved in family.Types)
            {
                var existing = Model.Types.FirstOrDefault(t => !t.IsCreate && t.SavedPath == family.Path && t.PersistedName == saved.PersistedName);
                if (existing == null)
                {
                    if (!Model.Types.Any(t => !t.IsCreate && t.Name == saved.Name))
                        Model.Types.Add(saved);
                }
                else if (!existing.IsDirty && saved.Configuration != null)
                {
                    existing.Configuration = saved.Configuration;
                    existing.MarkSaved();
                }
            }

            SelectType(Model.CurrentType);
        }

        internal void RememberOpenTypeIdentity(string path, string typeName)
        {
            if (Model.CurrentType == null)
                return;
            Model.CurrentType.Name = typeName;
            Model.CurrentType.PersistedName = typeName;
            Model.CurrentType.SavedPath = Model.CurrentType.SourcePath = path;
            Model.CurrentType.RequireSave();
        }

        internal void CompleteDeletedFamily(FamilyPackage family, FamilyTypeItem deleted)
        {
            if (deleted != null && Model.Types.Contains(deleted))
                RemoveTypeFromList(deleted);
            Model.WorkingFamilyPath = Model.LastSavedPath = family.Path;
            Model.WorkingSourcePath = family.SourcePath ?? family.Path;
            UpdateProjectHeader();
            Model.SectionCatalog = family.Catalog;
            foreach (var item in Model.Types.Where(t => !t.IsCreate))
            {
                item.SavedPath = family.Path;
                item.SourcePath = family.SourcePath ?? family.Path;
                var saved = family.Types.FirstOrDefault(t => t.PersistedName == item.PersistedName);
                if (!item.IsDirty && saved?.Configuration != null)
                {
                    item.Configuration = saved.Configuration;
                    item.MarkSaved();
                }
            }

            foreach (var saved in family.Types)
                if (!Model.Types.Any(t => !t.IsCreate && t.PersistedName == saved.PersistedName))
                    Model.Types.Add(saved);
            if (Model.CurrentType != null && Model.CurrentType.Configuration != null)
                SelectType(Model.CurrentType);
        }

        internal bool SwitchFamilyContext(string key)
        {
            if (Model.FamilyContextKey == key)
                return true;
            if (Model.FamilyContextKey != null)
                Model.Workspaces[Model.FamilyContextKey] = new OV_VentConfiguratorM.WorkspaceState
                {
                    Family = new FamilyPackage
                    {
                        Path = Model.WorkingFamilyPath,
                        SourcePath = Model.WorkingSourcePath,
                        Catalog = Model.SectionCatalog,
                        Types = Model.Types.Where(t => !t.IsCreate).ToList(),
                        UpdateNotice = Model.FamilyUpdateNotice
                    },
                    Selected = Model.CurrentType
                };
            Model.FamilyContextKey = key;
            OV_VentConfiguratorM.WorkspaceState workspace;
            if (Model.Workspaces.TryGetValue(key, out workspace))
            {
                SetFamily(workspace.Family);
                if (workspace.Selected != null && workspace.Selected.Configuration != null)
                {
                    Model.IsSynchronizing = true;
                    try
                    {
                        _mainWindow.SelectedType = workspace.Selected;
                    }
                    finally
                    {
                        Model.IsSynchronizing = false;
                    }

                    SelectType(workspace.Selected);
                }

                return true;
            }

            SetFamily(new FamilyPackage());
            return false;
        }

        internal string AvailableAutomaticName(string requested, IEnumerable<string> fileNames)
        {
            var names = fileNames.Concat(Model.Types.Where(t => !t.IsCreate && t != Model.CurrentType).Select(t => t.Name)).ToList();
            return FamilyTypeItem.AvailableAutomaticName(requested, names);
        }

        internal void MergeDiscoveredFamily(FamilyPackage family, string currentName)
        {
            if (Model.CurrentType != null && Model.CurrentType.Name != currentName)
            {
                Model.IsSynchronizing = true;
                try
                {
                    Model.CurrentType.Name = currentName;
                    _mainWindow.TypeName = currentName;
                }
                finally
                {
                    Model.IsSynchronizing = false;
                }
            }

            foreach (var saved in family.Types)
            {
                var existing = Model.Types.FirstOrDefault(t => !t.IsCreate && t.PersistedName == saved.PersistedName && (string.IsNullOrWhiteSpace(t.SavedPath) || string.Equals(t.SavedPath, saved.SavedPath, StringComparison.OrdinalIgnoreCase) || string.Equals(t.SourcePath, Model.WorkingSourcePath, StringComparison.OrdinalIgnoreCase)));
                if (existing == null)
                    Model.Types.Add(saved);
                else
                {
                    existing.SavedPath = saved.SavedPath;
                    existing.SourcePath = saved.SourcePath;
                    if (existing != Model.CurrentType && !existing.IsDirty && saved.Configuration != null)
                    {
                        existing.Configuration = saved.Configuration;
                        existing.MarkSaved();
                    }
                }
            }
        }

        internal void SetFamily(FamilyPackage family)
        {
            Model.FamilyUpdateNotice = family.UpdateNotice;
            Model.IsSynchronizing = true;
            try
            {
                foreach (var block in _observedSections)
                    block.PropertyChanged -= SectionChanged;
                _observedSections.Clear();
                ClearDimensionObservers();
                ClearInfoObserver();
                Model.Types.Clear();
                Model.CurrentType = null;
                Model.Configuration = null;
                Model.SelectedSection = null;
                Model.SectionCatalog = family.Catalog;
                Model.LastSavedPath = Model.WorkingFamilyPath = family.Path;
                Model.WorkingSourcePath = family.SourcePath ?? family.Path;
                UpdateProjectHeader();
                foreach (var item in family.Types)
                    Model.Types.Add(item);
                _mainWindow.ClearConfiguration();
            }
            finally
            {
                Model.IsSynchronizing = false;
            }

            var first = family.Types.FirstOrDefault(t => t.Configuration != null && t.PersistedName == family.SelectedTypeName) ?? family.Types.FirstOrDefault(t => t.Configuration != null);
            if (first == null)
            {
                SetBusy(Model.IsBusy);
                SetStatus("Выберите «Создать тип», чтобы настроить новую установку или выберите уже имеющийся тип.", false);
            }
            else
            {
                Model.IsSynchronizing = true;
                try
                {
                    _mainWindow.SelectedType = first;
                }
                finally
                {
                    Model.IsSynchronizing = false;
                }

                SelectType(first);
            }
        }
    }
}
