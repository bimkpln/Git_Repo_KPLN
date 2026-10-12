using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;

namespace KPLN_Tools_OVVK.Forms.ViewModels
{
    internal sealed partial class OV_VentConfiguratorVM
    {
        private readonly List<SectionValue> _observedSections = new List<SectionValue>();
        private readonly List<DimensionValue> _observedDimensions = new List<DimensionValue>();
        private InstallationInfo _observedInfo;
        private bool _updatingCalculation;
        private readonly List<BooleanValue> _observedBooleans = new List<BooleanValue>();

        private void SetConfiguration(InstallationConfiguration configuration)
        {
            Model.Configuration = configuration;
            ClearInfoObserver();
            _observedInfo = configuration.Info;
            _observedInfo.PropertyChanged += InfoChanged;
            int index = Math.Max(0, Math.Min(Model.CurrentType.SelectedSectionIndex, configuration.Blocks.Count - 1));
            Model.SelectedSection = configuration.Blocks[index];
            _mainWindow.RenderConfiguration(configuration);
            _mainWindow.SyncSectionControls();
            ObserveSections();
        }

        private void ObserveSections()
        {
            foreach (var block in _observedSections)
                block.PropertyChanged -= SectionChanged;
            _observedSections.Clear();
            foreach (var block in Model.Configuration.Blocks)
            {
                block.PropertyChanged += SectionChanged;
                _observedSections.Add(block);
            }

            ClearDimensionObservers();
            foreach (var dimension in Model.Configuration.Dimensions().Values)
            {
                dimension.PropertyChanged += DimensionChanged;
                _observedDimensions.Add(dimension);
            }

            foreach (var flag in Model.Configuration.SharedBooleans.Values)
            {
                flag.PropertyChanged += BooleanChanged;
                _observedBooleans.Add(flag);
            }
        }

        private void ClearDimensionObservers()
        {
            foreach (var dimension in _observedDimensions)
                dimension.PropertyChanged -= DimensionChanged;
            _observedDimensions.Clear();
            foreach (var flag in _observedBooleans)
                flag.PropertyChanged -= BooleanChanged;
            _observedBooleans.Clear();
        }

        private void ClearInfoObserver()
        {
            if (_observedInfo != null)
                _observedInfo.PropertyChanged -= InfoChanged;
            _observedInfo = null;
        }

        private void InfoChanged(object sender, PropertyChangedEventArgs e)
        {
            if (Model.IsSynchronizing || Model.CurrentType == null)
                return;
            if (e.PropertyName == nameof(InstallationInfo.SystemName) || e.PropertyName == nameof(InstallationInfo.Mark))
            {
                Model.CurrentType.UpdateAutomaticName(Model.Types.Where(t => !t.IsCreate && t != Model.CurrentType).Select(t => t.Name));
                Model.IsSynchronizing = true;
                try
                {
                    _mainWindow.TypeName = Model.CurrentType.Name;
                }
                finally
                {
                    Model.IsSynchronizing = false;
                }
            }

            Model.CurrentType.MarkChanged();
        }

        private void SectionChanged(object sender, PropertyChangedEventArgs e)
        {
            if (!Model.IsSynchronizing && Model.CurrentType != null)
                Model.CurrentType.MarkChanged();
            RecalculateLocally();
            if (!Model.IsSynchronizing && e.PropertyName == "Type")
            {
                ObserveSections();
                if (ReferenceEquals(sender, Model.SelectedSection))
                    _mainWindow.RenderSectionParameters();
            }
        }

        private void BooleanChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "Value" || _updatingCalculation || Model.IsSynchronizing || Model.CurrentType == null)
                return;
            Model.CurrentType.MarkChanged();
            RecalculateLocally();
        }

        private void DimensionChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Text" && !_updatingCalculation && !((DimensionValue)sender).IsCalculated)
            {
                if (!Model.IsSynchronizing && Model.CurrentType != null)
                    Model.CurrentType.MarkChanged();
                RecalculateLocally();
            }
        }

        private void RecalculateLocally()
        {
            if (Model.Configuration == null || Model.IsSynchronizing || _updatingCalculation)
                return;
            _updatingCalculation = true;
            try
            {
                Model.Configuration.RecalculateLocal();
                _mainWindow.RenderBlocks();
            }
            finally
            {
                _updatingCalculation = false;
            }
        }

        internal void ApplyCalculatedDimensions(InstallationConfiguration calculated)
        {
            if (Model.Configuration == null)
                return;
            var targets = Model.Configuration.Dimensions();
            _updatingCalculation = true;
            try
            {
                foreach (var pair in calculated.Dimensions())
                {
                    DimensionValue target;
                    if (!targets.TryGetValue(pair.Key, out target))
                        continue;
                    target.SetCalculatedMode(pair.Value.IsCalculated);
                    if (pair.Value.IsCalculated && pair.Value.IsCurrent)
                        target.ReadResult(pair.Value.Millimeters(pair.Key));
                }

                foreach (var pair in calculated.SharedBooleans)
                {
                    var target = Model.Configuration.SharedBooleans[pair.Key];
                    target.SetCalculatedMode(pair.Value.IsCalculated);
                    if (pair.Value.IsCalculated)
                        target.ReadResult(pair.Value.Value);
                }
            }
            finally
            {
                _updatingCalculation = false;
            }
        }

        internal void MoveSection(int index, int target)
        {
            EditSections(() => Model.Configuration.Move(index, target));
        }

        internal void InsertSection(int index)
        {
            EditSections(() => Model.Configuration.InsertAfter(index));
        }

        internal void RemoveSection(int index)
        {
            EditSections(() => Model.Configuration.Remove(index));
        }

        internal IList<SectionTypeChoice> GetSelectedSectionChoices()
        {
            int index = Model.Configuration.Blocks.IndexOf(Model.SelectedSection);
            var choices = Model.Configuration.Catalog[Model.Configuration.SlotAt(index)].Choices.ToList();
            var selected = Model.SelectedSection.Type;
            if (selected != null)
            {
                if (!choices.Any(t => t.Key == selected.Key))
                    choices.Insert(0, selected);
                Model.SelectedSection.Type = choices.Single(t => t.Key == selected.Key);
            }

            return choices;
        }

        private void EditSections(Action action)
        {
            if (Model.IsBusy || Model.Configuration == null || FamilyService.IsClosed)
                return;
            try
            {
                action();
                Model.CurrentType.MarkChanged();
                _mainWindow.SyncSectionControls();
                if (Model.SelectedSection == null || !Model.Configuration.Blocks.Contains(Model.SelectedSection))
                    Model.SelectedSection = Model.Configuration.Blocks[Model.Configuration.FirstSectionIndex];
                ObserveSections();
                _mainWindow.RenderBlocks();
                _mainWindow.RenderSelectedSection();
                SetBusy(false);
                SetStatus("Количество секций: " + Model.Configuration.SectionCount + ".", false);
                RecalculateLocally();
            }
            catch (Exception ex)
            {
                _mainWindow.SyncSectionControls();
                SetStatus(ex.Message, true);
            }
        }
    }
}
