using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace KPLN_Tools_Lib.Forms.Entities
{
    internal sealed class CreateSpacesM : INotifyPropertyChanged
    {
        private Phase _selectedPhase;

        internal CreateSpacesM(Document document, IEnumerable<RevitLinkInstance> links)
        {
            Links = links.OrderBy(link => link.Name)
                .Select(link => new CreateSpacesLinkM(link)).ToArray();
            Phases = document.Phases.Cast<Phase>().ToArray();

            ElementId viewPhaseId = document.ActiveView?.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId();
            SelectedPhase = Phases.FirstOrDefault(phase => phase.Id.Equals(viewPhaseId)) ?? Phases.LastOrDefault();
        }

        public CreateSpacesLinkM[] Links { get; }
        public Phase[] Phases { get; }

        public Phase SelectedPhase
        {
            get => _selectedPhase;
            set
            {
                _selectedPhase = value;
                NotifyPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool CanRun => SelectedPhase != null && Links.Any(link => link.IsSelected);

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    internal sealed class CreateSpacesLinkM : INotifyPropertyChanged
    {
        private bool _isSelected = true;

        internal CreateSpacesLinkM(RevitLinkInstance link)
        {
            RevitLink = link;
            Name = string.Format("{0} [id: {1}]", link.Name, link.Id);
        }

        internal RevitLinkInstance RevitLink { get; }
        public string Name { get; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
