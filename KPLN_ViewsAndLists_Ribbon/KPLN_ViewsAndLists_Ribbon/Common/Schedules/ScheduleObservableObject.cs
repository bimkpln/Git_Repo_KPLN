using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_ViewsAndLists_Ribbon.Common.Schedules
{
    public abstract class ScheduleObservableObject : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void Notify([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        protected void Change<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (Equals(field, value)) return;
            field = value;
            Notify(name);
        }
    }
}
