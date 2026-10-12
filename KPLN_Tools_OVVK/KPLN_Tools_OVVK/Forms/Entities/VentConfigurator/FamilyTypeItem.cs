using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class FamilyTypeItem : EditableValue
    {
        public bool IsCreate
        { get; internal set; }

        internal InstallationConfiguration Configuration;
        internal int SelectedSectionIndex, SelectedTab;
        internal string SavedPath, SourcePath, PersistedName, LoadError;
        internal bool HasAutomaticName;
        internal bool IsDirty = true;
        private bool _needsImportSave;
        private string _savedSignature;

        private string StateSignature()
        {
            string name = (Name ?? "").Trim();
            return name.Length + ":" + name + Configuration?.Signature();
        }

        internal void MarkSaved()
        {
            _savedSignature = StateSignature();
            _needsImportSave = false;
            IsDirty = false;
            Changed(nameof(NeedsReview));
            Changed(nameof(ToolTipText));
        }

        internal void MarkImported()
        {
            MarkSaved();
            _needsImportSave = true;
            IsDirty = true;
        }

        internal bool HasUserChanges
        {
            get
            {
                return _savedSignature == null || _savedSignature != StateSignature();
            }
        }

        internal void MarkChanged()
        {
            IsDirty = _needsImportSave || HasUserChanges;
            Changed(nameof(NeedsReview));
            Changed(nameof(ToolTipText));
        }

        internal void RequireSave()
        {
            _savedSignature = null;
            IsDirty = true;
        }

        internal void UpdateAutomaticName(IEnumerable<string> existingNames)
        {
            if (HasAutomaticName)
                Name = AvailableAutomaticName(Configuration.Info.AutomaticTypeName, existingNames);
        }

        private string _name;

        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
                Changed();
                Changed(nameof(DisplayName));
            }
        }

        public bool NeedsReview
        {
            get
            {
                return !IsCreate && (Configuration == null || Configuration.Problems().Count > 0);
            }
        }

        public string ToolTipText
        {
            get
            {
                return DisplayName;
            }
        }

        public string DisplayName
        {
            get
            {
                return IsCreate ? "Создать тип" : string.IsNullOrWhiteSpace(Name) ? "Без имени" : Name;
            }
        }

        internal static string AvailableAutomaticName(string requested, IEnumerable<string> existingNames)
        {
            var names = new HashSet<string>(existingNames.Select(n => (n ?? "").Trim()), StringComparer.OrdinalIgnoreCase);
            if (!names.Contains(requested))
                return requested;
            int number = 1;
            while (names.Contains(requested + " " + number))
                number++;
            return requested + " " + number;
        }

        internal FamilyTypeItem CreateCopy(IEnumerable<FamilyTypeItem> existing)
        {
            if (Configuration == null)
                throw new InvalidOperationException("Не удалось прочитать исходный тип для копирования.");
            var copy = new FamilyTypeItem
            {
                Name = AvailableAutomaticName(Name + " — Копия", existing.Select(t => t.Name)),
                Configuration = Configuration.Copy(),
                SavedPath = SavedPath,
                SourcePath = SourcePath,
                SelectedTab = SelectedTab,
                SelectedSectionIndex = SelectedSectionIndex,
                HasAutomaticName = false
            };
            // PersistedName пуст: копия создаётся новым типом, исходный не переименовывается.
            copy.Configuration.CopySourceTypeName = PersistedName ?? Configuration.CopySourceTypeName ?? Configuration.Catalog.SourceTypeName;
            return copy;
        }

        internal static FamilyTypeItem CreateDraft(IEnumerable<FamilyTypeItem> existing, SectionCatalog catalog)
        {
            return new FamilyTypeItem
            {
                Name = AvailableAutomaticName("Новый тип", existing.Where(t => !t.IsCreate).Select(t => t.Name)),
                HasAutomaticName = true,
                Configuration = InstallationConfiguration.CreateDraft(catalog, newTypeDefaults: true)
            };
        }
    }
}
