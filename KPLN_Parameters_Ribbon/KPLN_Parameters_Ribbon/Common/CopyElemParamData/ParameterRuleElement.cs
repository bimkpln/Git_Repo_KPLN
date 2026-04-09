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



        public static bool SaveData(string path, ObservableCollection<ParameterRuleElement> collection)
        {
            try
            {
                if (collection == null)
                    return false;

                List<string> dataparts = new List<string>();

                foreach (ParameterRuleElement el in collection)
                {
                    string[] parts = new string[3];

                    parts[0] = el != null && el.SelectedCategory != null
                        ? el.SelectedCategory.Name
                        : "null";

                    ListBoxElement sourceElement = el != null ? el.SelectedSourceParameter as ListBoxElement : null;
                    Parameter sourceParam = sourceElement != null ? sourceElement.Data as Parameter : null;

                    if (sourceParam != null && sourceParam.Definition != null)
                    {
                        parts[1] = sourceParam.Definition.Name;
                    }
                    else
                    {
                        parts[1] = "null";
                    }

                    ListBoxElement targetElement = el != null ? el.SelectedTargetParameter as ListBoxElement : null;
                    Parameter targetParam = targetElement != null ? targetElement.Data as Parameter : null;

                    if (targetParam != null && targetParam.Definition != null)
                    {
                        parts[2] = targetParam.Definition.Name;
                    }
                    else
                    {
                        parts[2] = "null";
                    }

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

            string fileText = File.ReadAllText(path);

            List<string> dataparts = fileText
                .Split(new string[] { Variables.separator_element }, StringSplitOptions.RemoveEmptyEntries)
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

                ParameterRuleElement rule = (parent.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>)[z];
                System.Windows.Forms.Application.DoEvents();

                bool categoryFound = false;
                bool sourceFound = false;
                bool targetFound = false;

                if (parts[0] != "null")
                {
                    foreach (ListBoxElement cat in rule.Categories)
                    {
                        if (cat.Name == parts[0])
                        {
                            rule.SelectedCategory = cat;
                            System.Windows.Forms.Application.DoEvents();
                            categoryFound = true;
                            break;
                        }
                    }
                }

                if (categoryFound && parts[1] != "null")
                {
                    foreach (ListBoxElement par in rule.SourceParameters)
                    {
                        Parameter param = par.Data as Parameter;
                        if (param == null)
                            continue;

                        if (param.Definition != null && param.Definition.Name == parts[1])
                        {
                            rule.SelectedSourceParameter = par;
                            System.Windows.Forms.Application.DoEvents();
                            sourceFound = true;
                            break;
                        }

                        if (param.Id != null && param.Id.ToString() == parts[1])
                        {
                            rule.SelectedSourceParameter = par;
                            System.Windows.Forms.Application.DoEvents();
                            sourceFound = true;
                            break;
                        }
                    }
                }

                if (categoryFound && parts[2] != "null")
                {
                    foreach (ListBoxElement par in rule.TargetParameters)
                    {
                        Parameter param = par.Data as Parameter;
                        if (param == null)
                            continue;

                        if (param.Definition != null && param.Definition.Name == parts[2])
                        {
                            rule.SelectedTargetParameter = par;
                            System.Windows.Forms.Application.DoEvents();
                            targetFound = true;
                            break;
                        }

                        if (param.Id != null && param.Id.ToString() == parts[2])
                        {
                            rule.SelectedTargetParameter = par;
                            System.Windows.Forms.Application.DoEvents();
                            targetFound = true;
                            break;
                        }
                    }
                }

                if (!categoryFound && parts[0] != "null")
                    Print(string.Format("[Категория не найдена:] <{0}>", parts[0]), MessageType.Error);

                if (parts[1] != "null" && !sourceFound)
                    Print(string.Format("[Параметр-источник не найден:] <{0}>", parts[1]), MessageType.Error);

                if (parts[2] != "null" && !targetFound)
                    Print(string.Format("[Параметр-назначение не найден:] <{0}>", parts[2]), MessageType.Error);
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
                TargetParameters = new ObservableCollection<ListBoxElement>();

                if (_selectedCategory != null && _selectedCategory.SubElements != null)
                {
                    foreach (ListBoxElement element in _selectedCategory.SubElements)
                    {
                        SourceParameters.Add(element);
                    }

                    foreach (ListBoxElement element in _selectedCategory.SubElements)
                    {
                        TargetParameters.Add(element);
                    }
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
