using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.Common.CopyElemParamData
{
    public class ParameterRuleElement : INotifyPropertyChanged
    {



        public static bool SaveData(string path, IEnumerable<ParameterRuleElement> collection)
        {
            try
            {
                if (collection == null)
                    return false;

                List<string> dataparts = new List<string>();

                foreach (ParameterRuleElement el in collection)
                {
                    if (el.IsCompletelyEmpty)
                        continue;

                    string[] parts = new string[3];
                    parts[0] = el.SelectedCategory?.Name ?? "null";
                    parts[1] = GetParameterName(el.SelectedSourceParameter) ?? "null";
                    parts[2] = GetParameterName(el.SelectedTargetParameter) ?? "null";

                    dataparts.Add(string.Join(Variables.separator_sub_element, parts));
                }

                File.WriteAllText(path, string.Join(Variables.separator_element, dataparts));
                return true;
            }
            catch (Exception e)
            {
                PrintError(e);
                return false;
            }
        }




        public static void LoadData(ParamSetter parent, string path)
        {
            parent.RulesControll.ItemsSource = null;
            parent.RulesControll.ItemsSource = new ObservableCollection<ParameterRuleElement>();

            List<string> dataparts = File.ReadAllText(path)
                .Split(new string[] { Variables.separator_element }, StringSplitOptions.None)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            for (int z = 0; z < dataparts.Count; z++)
            {
                List<string> parts = dataparts[z]
                    .Split(new string[] { Variables.separator_sub_element }, StringSplitOptions.None)
                    .ToList();

                if (parts.Count < 3)
                {
                    Print(string.Format("[Ошибка чтения строки настроек:] <{0}>", dataparts[z]), MessageType.Error);
                    continue;
                }

                parent.AddRule();

                List<string> parts = dataparts[z]
                    .Split(new string[] { Variables.separator_sub_element }, StringSplitOptions.None)
                    .ToList();

                while (parts.Count < 3)
                    parts.Add("null");

                ParameterRuleElement rule = (parent.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>)[z];
                System.Windows.Forms.Application.DoEvents();

                bool category_found = parts[0] == "null";
                bool source_found = parts[1] == "null";
                bool target_found = parts[2] == "null";

                if (parts[0] != "null")
                {
                    foreach (ListBoxElement cat in rule.Categories)
                    {
                        rule.SelectedCategory = cat;
                        System.Windows.Forms.Application.DoEvents();
                        category_found = true;
                        break;
                    }
                }

                if (rule.SourceParameters != null)
                {
                    foreach (ListBoxElement par in rule.SourceParameters)
                    {
                        if ((par.Data as Parameter).Definition.Name == parts[1] ||
                            (par.Data as Parameter).Id.ToString() == parts[1])
                        {
                            rule.SelectedSourceParameter = par;
                            System.Windows.Forms.Application.DoEvents();
                            source_found = true;
                            break;
                        }
                    }
                }

                if (rule.TargetParameters != null)
                {
                    foreach (ListBoxElement par in rule.TargetParameters)
                    {
                        if ((par.Data as Parameter).Definition.Name == parts[2] ||
                            (par.Data as Parameter).Id.ToString() == parts[2])
                        {
                            rule.SelectedTargetParameter = par;
                            System.Windows.Forms.Application.DoEvents();
                            target_found = true;
                            break;
                        }
                    }
                }

                if (!category_found)
                    Print(string.Format("[Категория не найдена:] <{0}>", parts[0]), MessageType.Error);

                if (!source_found)
                    Print(string.Format("[Параметр не найден:] <{0}>", parts[1]), MessageType.Error);

                if (!target_found)
                    Print(string.Format("[Параметр не найден:] <{0}>", parts[2]), MessageType.Error);
            }

            parent.UpdateRunEnability();
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
                TargetParameters = new ObservableCollection<ListBoxElement>();

                SelectedSourceParameter = null;
                SelectedTargetParameter = null;

                if (_selectedCategory != null)
                {
                    foreach (ListBoxElement element in _selectedCategory.SubElements)
                        SourceParameters.Add(element);

                    foreach (ListBoxElement element in _selectedCategory.SubElements)
                        TargetParameters.Add(element);
                }

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


        public bool IsCompletelyEmpty =>
            SelectedCategory == null &&
            SelectedSourceParameter == null &&
            SelectedTargetParameter == null;

        public bool IsPartiallyFilled =>
            !IsCompletelyEmpty &&
            (SelectedCategory == null ||
             SelectedSourceParameter == null ||
             SelectedTargetParameter == null);

        public static string GetParameterName(ListBoxElement element)
        {
            if (element?.Data is Parameter p)
                return p.Definition.Name;
            return null;
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
