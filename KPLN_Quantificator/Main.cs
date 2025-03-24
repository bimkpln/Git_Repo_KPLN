using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Interop;
using Autodesk.Navisworks.Api.Plugins;
using KPLN_Quantificator.Forms;
using KPLN_Quantificator.Services;
using System;


namespace KPLN_Quantificator
{
    [Plugin("KPLN Extention", "KPLN", DisplayName = "Quantification")]
    [RibbonLayout("KPLN_Quantificator.xaml")]
    [RibbonTab("KPLN_Tab", DisplayName = "KPLN Extention")]
    [Command("ID_Button_A", DisplayName = "Поисковые наборы", Icon = "Source\\scope_set_small.png", LargeIcon = "Source\\scope_set_big.png", ToolTip = "Создание поисковых наборов по выбранному параметру WBS", CanToggle = true)]
    [Command("ID_Button_B", DisplayName = "Добавить элементы", Icon = "Source\\update_q_small.png", LargeIcon = "Source\\update_q_big.png", ToolTip = "Добавление элементов (в которые непосредствунно будут добавляться объекты модели) в существующую структуру книги Quantification ", CanToggle = true)]
    [Command("ID_Button_C", DisplayName = "Добавить объекты", Icon = "Source\\create_items_small.png", LargeIcon = "Source\\create_items_big.png", ToolTip = "Наполнение каталогов Quantification объектами модели из выбранных поисковых наборов", CanToggle = true)]
    [Command("ID_Button_D", DisplayName = "Добавить ресурсы", Icon = "Source\\match_resources_small.png", LargeIcon = "Source\\match_resources_big.png", ToolTip = "Сопоставление ресурсов с элементами по выбранному параметру RBS", CanToggle = true)]
    [Command("ID_Button_E", DisplayName = "Сгруппировать коллизии", Icon = "Source\\group_c_small.png", LargeIcon = "Source\\group_c_big.png", ToolTip = "Группировка коллизий по выбранным параметрам. Сделано на основе «Group Clashes». Горячие клавиши - Shift + G", CanToggle = true)]
    [Command("ID_Button_F", DisplayName = "Подсчет коллизий", Icon = "Source\\counter_small.png", LargeIcon = "Source\\counter_big.png", ToolTip = "Подсчет количества коллизий по разделам (раздел выделяется из имени)", CanToggle = true)]
    [Command("ID_Button_G", DisplayName = "Автоматический комментарий", Icon = "Source\\comment_small.png", LargeIcon = "Source\\comment_big.png", ToolTip = "Создание текстового комментария. Для создания комментария в автоматическом режиме необходимо выделить элемент/элементы и нажать клавишу E", CanToggle = true)]
    [Command("ID_Button_H", DisplayName = "Настройка для пакетного переименования точек обзора", Icon = "Source\\rename_small.png", LargeIcon = "Source\\rename_big.png", ToolTip = "Настройка для пакетного переименования точек обзора.\nДля переименования точки обзора - задайте параметры в данном окне, после чего выберите необходимую точку обзора и нажмите клавишу Q", CanToggle = true)]
    public class Main : CommandHandlerPlugin
    {
        public override int ExecuteCommand(string name, params string[] parameters)
        {
            if (GlobalPreferences.state == 0)
            {
                GlobalPreferences.state = 1;

                try
                {
                    switch (name)
                    {
                        case "ID_Button_A":
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument.Models.Count != 0)
                                {
                                    CreateSelectionSetsForm form1 = new CreateSelectionSetsForm();
                                    form1.Show();
                                }
                                else { GlobalPreferences.state = 0; }

                                break;
                            }
                        case "ID_Button_B":
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument.SelectionSets.RootItem.Children.Count != 0)
                                {
                                    ElemtsToQuantItemsForm form2 = new ElemtsToQuantItemsForm();
                                    form2.Show();
                                }
                                else { GlobalPreferences.state = 0; }

                                break;
                            }
                        case "ID_Button_C":
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument.SelectionSets.RootItem.Children.Count != 0)
                                {
                                    CreateQuantItemsForm form3 = new CreateQuantItemsForm();
                                    form3.Show();
                                }
                                else { GlobalPreferences.state = 0; }

                                break;
                            }
                        case "ID_Button_D":
                            {
                                GlobalPreferences.Update();
                                ElementsToResourcesCompareForm form4 = new ElementsToResourcesCompareForm();
                                form4.Show();
                                GlobalPreferences.state = 0;

                                break;
                            }
                        case "ID_Button_E":
                            {
                                ClashGroupsForm clashGroupsForm = new ClashGroupsForm();
                                try
                                {
                                    clashGroupsForm.SearchText.Text = ClashCurrentIssue.CurrentTest?.DisplayName;
                                }
                                catch (NullReferenceException) { }
                                clashGroupsForm.ShowDialog();
                                GlobalPreferences.state = 0;

                                break;
                            }
                        case "ID_Button_F":
                            {
                                ClashesCounter.Prepare();
                                ClashesCounter.Execute();
                                ClashesCounter.PrintResult();
                                GlobalPreferences.state = 0;

                                break;
                            }
                        case "ID_Button_G":
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument.CurrentSelection.SelectedItems.Count >= 1 && Autodesk.Navisworks.Api.Application.ActiveDocument.CurrentSelection.SelectedItems.Count <= 20)
                                {
                                    AddComment.GettingDataForAComment();
                                    AddComment.CreateViewpoint();
                                }
                                GlobalPreferences.state = 0;
                                break;
                            }
                        case "ID_Button_H":
                            {
                                if (Autodesk.Navisworks.Api.Application.ActiveDocument?.SavedViewpoints?.CurrentSavedViewpoint?.DisplayName != null)
                                {
                                    RenameViewForm renameViewForm = new RenameViewForm();
                                    renameViewForm.ShowDialog();
                                }
                                GlobalPreferences.state = 0;
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

                catch (Exception e)
                {
                    Output.PrintError(e);
                    GlobalPreferences.state = 0;

                    return 0;
                }
            }      
            
            return 0;
        }
    }
}


namespace KPLN_Quantificator_inputPlugin
{
    [Plugin("KPLN Extention_inputPlugin", "KPLN")]
    public class Main : InputPlugin
    {
        public override bool KeyUp(Autodesk.Navisworks.Api.View view, KeyModifiers modifier, ushort key, double timeOffset)
        {
            if (modifier == KeyModifiers.Shift && key == 71)
            {
                ClashGroupsForm clashGroupsForm = new ClashGroupsForm();
                try
                {
                    clashGroupsForm.SearchText.Text = ClashCurrentIssue.CurrentTest?.DisplayName;
                }
                catch (NullReferenceException) { }

                clashGroupsForm.ShowDialog();
                return true;
            }

            if (Autodesk.Navisworks.Api.Application.ActiveDocument.CurrentSelection.SelectedItems.Count >= 1 && Autodesk.Navisworks.Api.Application.ActiveDocument.CurrentSelection.SelectedItems.Count <= 20 && key == 69)
            {
                AddComment.GettingDataForAComment();
                AddComment.CreateViewpoint();
                return true;
            }

            if (Autodesk.Navisworks.Api.Application.ActiveDocument?.SavedViewpoints?.CurrentSavedViewpoint?.DisplayName != null && key == 81)
            {
                RenameViewForm renameViewForm = new RenameViewForm();
                renameViewForm.RenameViewPointObj();
            }

            return base.KeyUp(view, modifier, key, timeOffset);
        }
    }
}