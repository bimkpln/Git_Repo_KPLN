using System;
using System.Threading;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Data;
using Autodesk.Navisworks.Api.Takeoff;
using KPLN_Quantificator.Forms;

namespace KPLN_Quantificator
{
    [Plugin("KPLN Extention", "KPLN", DisplayName = "Quantification")]
    [RibbonLayout("KPLN_Quantificator.xaml")]
    [RibbonTab("KPLN_Tab", DisplayName = "KPLN Extention")]
    [Command("ID_Button_A", DisplayName = "Поисковые наборы", Icon = "Source\\scope_set_small.png", LargeIcon = "Source\\scope_set_big.png", ToolTip = "Создание поисковых наборов по выбранному параметру WBS", CanToggle = true)]
    [Command("ID_Button_B", DisplayName = "Добавить элементы", Icon = "Source\\update_q_small.png", LargeIcon = "Source\\update_q_big.png", ToolTip = "Добавление элементов (в которые непосредствунно будут добавляться объекты модели) в существующую структуру книги Quantification ", CanToggle = true)]
    [Command("ID_Button_C", DisplayName = "Добавить объекты", Icon = "Source\\create_items_small.png", LargeIcon = "Source\\create_items_big.png", ToolTip = "Наполнение каталогов Quantification объектами модели из выбранных поисковых наборов", CanToggle = true)]
    [Command("ID_Button_D", DisplayName = "Добавить ресурсы", Icon = "Source\\match_resources_small.png", LargeIcon = "Source\\match_resources_big.png", ToolTip = "Сопоставление ресурсов с элементами по выбранному параметру RBS", CanToggle = true)]
    [Command("ID_Button_E", DisplayName = "Сгруппировать коллизии", Icon = "Source\\group_c_small.png", LargeIcon = "Source\\group_c_big.png", ToolTip = "Группировка коллизий по выбранным параметрам. Сделано на основе «Group Clashes»", CanToggle = true)]
    public class Main : CommandHandlerPlugin
    {
        public override int ExecuteCommand(string name, params string[] parameters)
        {
            if (GlobalPreferences.state == 0)
            {
                GlobalPreferences.state = 1;
                switch (name)
                {
                    case "ID_Button_A":
                        {
                            try
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument.Models.Count != 0)
                                {
                                    CreateSelectionSetsForm form1 = new CreateSelectionSetsForm();
                                    form1.Show();
                                }
                                else { GlobalPreferences.state = 0; }
                            }
                            catch (Exception e)
                            {
                                Output.PrintError(e);
                                GlobalPreferences.state = 0;
                            }
                            break;
                        }
                    case "ID_Button_B":
                        {
                            try
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument.SelectionSets.RootItem.Children.Count != 0)
                                {
                                    ElemtsToQuantItemsForm form2 = new ElemtsToQuantItemsForm();
                                    form2.Show();
                                }
                                else { GlobalPreferences.state = 0; }
                            }
                            catch (Exception e)
                            {
                                Output.PrintError(e);
                                GlobalPreferences.state = 0;
                            }
                            break;
                        }
                    case "ID_Button_C":
                        {
                            try
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument.SelectionSets.RootItem.Children.Count != 0)
                                {
                                    CreateQuantItemsForm form3 = new CreateQuantItemsForm();
                                    form3.Show();
                                }
                                else { GlobalPreferences.state = 0; }
                            }
                            catch (Exception e)
                            {
                                Output.PrintError(e);
                                GlobalPreferences.state = 0;
                            }
                            break;
                        }
                    case "ID_Button_D":
                        {
                            try
                            {
                                GlobalPreferences.Update();
                                ElementsToResourcesCompareForm form4 = new ElementsToResourcesCompareForm();
                                form4.Show();
                                GlobalPreferences.state = 0;
                            }
                            catch (Exception e)
                            {
                                Output.PrintError(e);
                                GlobalPreferences.state = 0;
                            }
                            break;
                        }
                    case "ID_Button_E":
                        {
                            try
                            {
                                ClashGroupsForm form = new ClashGroupsForm();
                                form.ShowDialog();
                                GlobalPreferences.state = 0;
                            }
                            catch (Exception e)
                            {
                                Output.PrintError(e);
                                GlobalPreferences.state = 0;
                            }
                            break;
                        }
                    default:
                        {
                            GlobalPreferences.state = 0;
                            break;
                        }   
                }
                return 1;
            }
            return 0;
        }
    }
}

