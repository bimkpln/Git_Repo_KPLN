using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Publication.Common.Filters
{
    public class WPFFilterElement : INotifyPropertyChanged
    {
        public Guid Guid = Guid.NewGuid();
        private string _number;
        public string Number
        {
            get
            {
                if (_number == null) { return "-"; }
                return _number;
            }
            set
            {
                _number = value;
                NotifyPropertyChanged();
            }
        }
        private ObservableCollection<ListBoxParameter> _Parameters = new ObservableCollection<ListBoxParameter>();
        public ObservableCollection<ListBoxParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
            set 
            {
                _Parameters = value;
                NotifyPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            //Print(string.Format("NotifyPropertyChanged (" + propertyName + ")"), KPLN_Loader.Preferences.MessageType.System_Regular);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private ListBoxParameter _selectedParameter;
        public bool ElementPassesFilter(View view)
        {
            if (SelectedParameter.Parameter == null) { return true; }
            if (SelectedValue != null)
            {
                foreach (Parameter p in view.Parameters)
                {
                    if (p.Id.Equals(SelectedParameter.Parameter.Id))
                    {
                        string v = p.AsString();
                        if (v == null)
                        {
                            if (SelectedValue == string.Empty)
                            {
                                return true;
                            }
                            return false;
                        }
                        switch (SelectedType)
                        {
                            case "Равно":
                                if (v == SelectedValue) { return true; }
                                else { return false; }
                            case "Не равно":
                                if (v != SelectedValue) { return true; }
                                else { return false; }
                            case "Содержит":
                                if (v.Contains(SelectedValue)) { return true; }
                                else { return false; }
                            case "Не содержит":
                                if (!v.Contains(SelectedValue)) { return true; }
                                else { return false; }
                            case "Начиается с":
                                if (v.StartsWith(SelectedValue)) { return true; }
                                else { return false; }
                            case "Заканчивается на":
                                if (v.EndsWith(SelectedValue)) { return true; }
                                else { return false; }
                            default:
                                Print("Тип равенства неопределен!", MessageType.Critical);
                                return true;
                        }
                    }
                }
            }
            return true;
        }
        public ListBoxParameter SelectedParameter
        {
            get
            {
                return _selectedParameter;
            }
            set
            {
                _selectedParameter = value;
                NotifyPropertyChanged();
            }
        }
        private string _selectedType;
        public string SelectedType
        {
            get
            {
                return _selectedType;
            }
            set
            {
                _selectedType = value;
                NotifyPropertyChanged();
            }
        }
        private ObservableCollection<string> _types;
        public ObservableCollection<string> Types
        {
            get
            {
                return _types;
            }
            set
            {
                _types = value;
                NotifyPropertyChanged();
            }
        }
        private string _selectedValue;
        public string SelectedValue
        {
            get
            {
                return _selectedValue;
            }
            set
            {
                _selectedValue = value;
                NotifyPropertyChanged();
            }
        }
        private ObservableCollection<string> _values;
        public ObservableCollection<string> Values
        {
            get
            {
                return _values;
            }
            set
            {
                _values = value;
                NotifyPropertyChanged();
            }
        }
        public WPFFilterElement()
        { 
            //Print("WPFFilterElement", KPLN_Loader.Preferences.MessageType.System_OK); 
        }
        public void SetParameters(ObservableCollection<ListBoxParameter> parameters)
        {
            //Print("SetParameters", KPLN_Loader.Preferences.MessageType.System_Regular);
            Parameters = new ObservableCollection<ListBoxParameter>();
            foreach (var i in parameters)
            {
                Parameters.Add(i);
            }
        }
        public void SetTypes(ObservableCollection<string> values)
        {
            //Print("SetTypes", KPLN_Loader.Preferences.MessageType.System_Regular);
            Types = new ObservableCollection<string>();
            foreach (var i in values)
            {
                Types.Add(i);
            }
        }
    }
}
