using KPLN_Clashes_Ribbon.Core.Reports;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace KPLN_Clashes_Ribbon.Core
{
    /// <summary>
    /// Коллекция кнопок основных подотделов КПЛН
    /// </summary>
    public sealed class SubDepartmentBtn : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private Brush _delegateBtnBackground = Brushes.Transparent;

        public SubDepartmentBtn(int id, string name, string description = null)
        {
            Id = id;
            Name = name;
            Description = description;
        }

        public int Id { get; private set; }

        public string Name { get; private set; }

        public string Description { get; private set; }

        public Brush DelegateBtnBackground
        {
            get { return _delegateBtnBackground; }
            set 
            { 
                _delegateBtnBackground = value;
                NotifyPropertyChanged();
            }
        }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
