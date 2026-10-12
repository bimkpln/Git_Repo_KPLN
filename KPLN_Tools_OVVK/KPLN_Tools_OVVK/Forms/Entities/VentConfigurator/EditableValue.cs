using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal abstract class EditableValue : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void Changed([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
