using KPLN_Loader.Core.Entities;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Loader.Forms.Common
{
    public class WPFUser : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private string _name;
        private string _surname;
        private string _company;
        private SubDepartment _subDepartment;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }

        public string Surname
        {
            get => _surname;
            set
            {
                _surname = value;
                NotifyPropertyChanged();
            }
        }

        public string Company
        {
            get => _company;
            set
            {
                _company = value;
                NotifyPropertyChanged();
            }
        }

        public SubDepartment SubDepartment
        {
            get => _subDepartment;
            set
            {
                _subDepartment = value;
                NotifyPropertyChanged();
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
