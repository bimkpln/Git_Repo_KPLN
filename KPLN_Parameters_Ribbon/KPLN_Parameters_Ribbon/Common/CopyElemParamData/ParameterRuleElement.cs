using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.CopyElemParamData
{
    public class ParameterRuleElement : INotifyPropertyChanged
    {
        public static bool SaveData(string path, ObservableCollection<ParameterRuleElement> collection)
        {
            try
            {
                List<string> dataparts = new List<string>();
                foreach (ParameterRuleElement el in collection)
                {
                    string[] parts = new string[3];
                    parts[0] = el.SelectedCategory.Name;
                    if (el.SelectedSourceParameter == null)
                    {
                        parts[1] = "null";
                    }
                    else
                    {
                        parts[1] = ((el.SelectedSourceParameter as ListBoxElement).Data as Parameter).Definition.Name;
                    }
                    if (el.SelectedTargetParameter == null)
                    {
                        parts[2] = "null";
                    }
                    else
                    {
                        parts[2] = ((el.SelectedTargetParameter as ListBoxElement).Data as Parameter).Definition.Name;
                    }
                    dataparts.Add(string.Join(Variables.separator_sub_element, parts));
                }
                File.WriteAllText(path, string.Join(Variables.separator_element, dataparts));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public static void LoadData(ParamSetter parent, string path)
        {
            parent.RulesControll.ItemsSource = null;
            parent.RulesControll.ItemsSource = new ObservableCollection<ParameterRuleElement>();
            List<string> dataparts = File.ReadAllText(path).Split(new string[] { Variables.separator_element }, StringSplitOptions.RemoveEmptyEntries).ToList();
            for (int z = 0; z < dataparts.Count; z++)
            {
                parent.AddRule();
                System.Windows.Forms.Application.DoEvents();
                List<string> parts = dataparts[z].Split(new string[] { Variables.separator_sub_element }, StringSplitOptions.RemoveEmptyEntries).ToList();
                ParameterRuleElement rule = (parent.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>)[z];
                foreach (ListBoxElement cat in rule.Categories)
                {
                    if (cat.Name == parts[0])
                    {
                        rule.SelectedCategory = cat;
                        System.Windows.Forms.Application.DoEvents();
                        break;
                    }
                }
                foreach (ListBoxElement par in rule.SourceParameters)
                {

                    if ((par.Data as Parameter).Definition.Name == parts[1])
                    {
                        rule.SelectedSourceParameter = par;
                        System.Windows.Forms.Application.DoEvents();
                        break;
                    }
                    if ((par.Data as Parameter).Id.ToString() == parts[1])
                    {
                        rule.SelectedSourceParameter = par;
                        System.Windows.Forms.Application.DoEvents();
                        break;
                    }
                }
                foreach (ListBoxElement par in rule.TargetParameters)
                {
                    if ((par.Data as Parameter).Definition.Name == parts[2])
                    {
                        rule.SelectedTargetParameter = par;
                        System.Windows.Forms.Application.DoEvents();
                        break;
                    }
                    if ((par.Data as Parameter).Id.ToString() == parts[2])
                    {
                        rule.SelectedTargetParameter = par;
                        System.Windows.Forms.Application.DoEvents();
                        break;
                    }
                }
            }
        }
        public Guid Guid = Guid.NewGuid();
        private ObservableCollection<ListBoxElement> _categories = new ObservableCollection<ListBoxElement>();
        public ObservableCollection<ListBoxElement> Categories
        {
            get
            {
                return _categories;
            }
            set
            {
                _categories = value;
                NotifyPropertyChanged();
            }
        }
        private ObservableCollection<ListBoxElement> _sourceParameters = new ObservableCollection<ListBoxElement>();
        public ObservableCollection<ListBoxElement> SourceParameters
        {
            get
            {
                return _sourceParameters;
            }
            set
            {
                _sourceParameters = value;
                NotifyPropertyChanged();
            }
        }
        private ObservableCollection<ListBoxElement> _targetParameters = new ObservableCollection<ListBoxElement>();
        public ObservableCollection<ListBoxElement> TargetParameters
        {
            get
            {
                return _targetParameters;
            }
            set
            {
                _targetParameters = value;
                NotifyPropertyChanged();
            }
        }
        private ListBoxElement _selectedCategory;
        public ListBoxElement SelectedCategory
        {
            get
            {
                return _selectedCategory;
            }
            set
            {
                _selectedCategory = value;
                SourceParameters = new ObservableCollection<ListBoxElement>();
                foreach (ListBoxElement element in _selectedCategory.SubElements) { SourceParameters.Add(element); }
                TargetParameters = new ObservableCollection<ListBoxElement>();
                foreach (ListBoxElement element in _selectedCategory.SubElements) { TargetParameters.Add(element); }
                NotifyPropertyChanged();
            }
        }
        private ListBoxElement _selectedSourceParameter;
        public ListBoxElement SelectedSourceParameter
        {
            get
            {
                return _selectedSourceParameter;
            }
            set
            {
                _selectedSourceParameter = value;
                NotifyPropertyChanged();
            }
        }
        private ListBoxElement _selectedTargetParameter;
        public ListBoxElement SelectedTargetParameter
        {
            get
            {
                return _selectedTargetParameter;
            }
            set
            {
                _selectedTargetParameter = value;
                NotifyPropertyChanged();
            }
        }
        public ParameterRuleElement(ObservableCollection<ListBoxElement> categories)
        {
            Categories = categories;
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
