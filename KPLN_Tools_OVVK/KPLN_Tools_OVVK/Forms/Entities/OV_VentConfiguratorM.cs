using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;

namespace KPLN_Tools_OVVK.Forms.Entities
{
    internal sealed class OV_VentConfiguratorM : EditableValue
    {
        internal sealed class WorkspaceState
        {
            internal FamilyPackage Family;
            internal FamilyTypeItem Selected;
        }

        private bool _isBusy;

        public bool IsBusy
        {
            get
            {
                return _isBusy;
            }

            internal set
            {
                _isBusy = value;
                RefreshState();
            }
        }

        public bool CanEdit
        {
            get
            {
                return !IsBusy;
            }
        }

        public bool HasConfiguration
        {
            get
            {
                return CanEdit && Configuration != null;
            }
        }

        public bool HasCurrentType
        {
            get
            {
                return CanEdit && CurrentType != null;
            }
        }

        public bool CanUpdateFamily
        {
            get
            {
                return CanEdit && CanUpdate && !string.IsNullOrWhiteSpace(WorkingFamilyPath);
            }
        }

        public bool CanUpdate
        { get; internal set; }

        public string ProjectHeader
        { get; internal set; }

        public string StatusText
        { get; internal set; }

        public string StatusTooltip
        { get; internal set; }

        public void RefreshState()
        {
            Changed(string.Empty);
        }

        internal string LastSavedPath;
        internal string WorkingFamilyPath;
        internal string WorkingSourcePath;

        public string FamilyUpdateNotice
        { get; internal set; }

        internal string FamilyContextKey;
        internal ProjectItem HeaderProject;
        internal readonly Dictionary<string, WorkspaceState> Workspaces = new Dictionary<string, WorkspaceState>(StringComparer.OrdinalIgnoreCase);
        internal string ErrorDetails;

        public bool HasOperationError
        { get; internal set; }

        internal readonly ObservableCollection<FamilyTypeItem> Types = new ObservableCollection<FamilyTypeItem>();
        internal FamilyTypeItem CurrentType;
        internal SectionCatalog SectionCatalog;
        internal bool IsSynchronizing;
        internal InstallationConfiguration Configuration;
        internal SectionValue SelectedSection;
    }
}
