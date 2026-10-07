using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private static TextBlock CategoryLabel(string text)
        {return new TextBlock{Text=text,Margin=new Thickness(0,8,0,5),TextWrapping=TextWrapping.Wrap};}
        private void AutoSaveCategoryAssignments()
        {
            if(!TEP.Engine.CanManageCategoryDictionary)return;
            try
            {
                Engine.LoadClassificationDictionary();
                if(Engine.DictionaryStatus.StartsWith("Словарь недоступен"))throw new InvalidOperationException(Engine.DictionaryStatus);
                var draft=Engine.CategorySettingsDraft();Engine.PublishCategoryDictionary(draft,Engine.DictionaryRevision);Engine.ApplyCategorySettings(draft);
            }
            catch(Exception ex){Status.Text="Распределение применено к расчёту, но общий словарь не обновлён: "+ex.Message;MessageBox.Show(this,Status.Text,"KPLN | Общий словарь",MessageBoxButton.OK,MessageBoxImage.Warning);}
        }
        private void CategorySettings_Click(object sender,RoutedEventArgs e)
        {
            Try(()=>
            {
                var window=BuildCategorySettingsWindow();window.Owner=this;
                window.ShowDialog();
                if(Engine.RequestedDetail!=null){CloseNavigationDialog(this);return;}
                // Parameter writes are explicit actions and survive closing the settings dialog.
                checkedConfiguration=null;
                SetBusy(true);
                try{departmentGroups=Engine.ScanDepartments(ReportProgress);RefreshDepartmentCards();}
                finally{SetBusy(false);}
            });
        }
        private Window BuildCategorySettingsWindow()
        {
            bool manager=TEP.Engine.CanManageCategoryDictionary;
            Engine.LoadClassificationDictionary();
            var draft=Engine.CategorySettingsDraft();
            string revision=Engine.DictionaryRevision,readError=Engine.DictionaryStatus.StartsWith("Словарь недоступен")?Engine.DictionaryStatus:null;
            var window=new Window{Title="KPLN | Настройка категорий",Width=1230,Height=880,MinWidth=1000,MinHeight=720,
                WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=Background,FontFamily=FontFamily,FontSize=14};
            window.Resources.MergedDictionaries.Add(Resources);
            var root=new Grid{Margin=new Thickness(18)};
            root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});
            root.RowDefinitions.Add(new RowDefinition{Height=new GridLength(1,GridUnitType.Star)});
            root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});
            var header=new StackPanel();header.Children.Add(new TextBlock{Text="Категория расчёта",FontWeight=FontWeights.SemiBold,FontSize=17,Margin=new Thickness(0,0,0,8)});
            var roles=DepartmentRoles.Where(r=>r.Key!="unknown").ToList();
            var category=new ComboBox{ItemsSource=roles,DisplayMemberPath="Label",SelectedValuePath="Key",ToolTip="Выберите категорию. Ниже задаются её источник и правила; они применяются ко всем показателям, использующим эту категорию."};
            header.Children.Add(category);root.Children.Add(header);
            var zones=new Grid{Margin=new Thickness(0,14,0,0)};
            zones.ColumnDefinitions.Add(new ColumnDefinition());zones.ColumnDefinitions.Add(new ColumnDefinition{Width=new GridLength(16)});zones.ColumnDefinitions.Add(new ColumnDefinition());
            Grid.SetRow(zones,1);root.Children.Add(zones);
            var roomContent=new StackPanel{Margin=new Thickness(2,2,8,2)};var familyContent=new StackPanel{Margin=new Thickness(2,2,8,2)};
            var roomZone=new GroupBox{Header="Помещения",Content=new ScrollViewer{Content=roomContent,VerticalScrollBarVisibility=ScrollBarVisibility.Auto}};
            var familyZone=new GroupBox{Header="Семейства",Content=new ScrollViewer{Content=familyContent,VerticalScrollBarVisibility=ScrollBarVisibility.Auto}};
            zones.Children.Add(roomZone);Grid.SetColumn(familyZone,2);zones.Children.Add(familyZone);
            var roomSource=new RadioButton{Content="Считать по помещениям",GroupName="CategoryInput",Margin=new Thickness(0,0,0,12),ToolTip="Площадь рассчитывается по контурам помещений, распределённых в эту категорию. Источник значений всегда системный параметр Назначение."};
            var familySource=new RadioButton{Content="Считать по семействам",GroupName="CategoryInput",Margin=new Thickness(0,0,0,12),ToolTip="Для этой категории используются выбранные типы семейств. Распределённые сюда помещения исключаются, чтобы не учитывать одну категорию дважды."};
            roomContent.Children.Add(roomSource);familyContent.Children.Add(familySource);
            roomContent.Children.Add(CategoryLabel("Значения для классификации"));
            roomContent.Children.Add(new TextBlock{Text="Н - Назначение. К - Имя комнаты. У нового тега включены оба варианта. Точное имя комнаты уточняет общее назначение.",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,8)});
            var tags=new CategoryTagEditor{Tag="classification-tags"};roomContent.Children.Add(tags);
            var apartmentOnly=new CheckBox{Content="Применять правила только к помещениям квартир",Margin=new Thickness(0,12,0,5),ToolTip="Назначение равно Квартира или заполнен КВ_Номер (кроме 0 и -). Ограничение не даёт отнести санузел офиса к площади квартир."};roomContent.Children.Add(apartmentOnly);
            roomContent.Children.Add(new TextBlock{Text="Квартирное помещение: «Назначение» = «Квартира» или заполнен «КВ_Номер».",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,8)});
            var writeRooms=new Button{Content="Запись параметров квартир",HorizontalAlignment=HorizontalAlignment.Left,ToolTip="Показывает все распределённые помещения для записи названий их категорий в выбранный параметр."};
            familyContent.Children.Add(CategoryLabel("Тип семейства из модели"));
            var picker=new ComboBox{ItemsSource=Engine.FamilyChoices,DisplayMemberPath="Display",IsTextSearchEnabled=true,ToolTip="Выберите тип семейства. В категорию попадут все его экземпляры в выбранных источниках и стадии."};familyContent.Children.Add(picker);
            var actions=new WrapPanel{Margin=new Thickness(0,8,0,8)};
            var add=new Button{Content="Добавить тип",ToolTip="Добавляет тип в эту категорию. Один тип можно назначить только одной категории."};
            var remove=new Button{Content="Убрать тип",ToolTip="Удаляет выбранную строку настройки. Элементы модели не удаляются."};actions.Children.Add(add);actions.Children.Add(remove);familyContent.Children.Add(actions);
            var families=new DataGrid{Height=260,SelectionMode=DataGridSelectionMode.Single};
            TextColumn(families,"Тип семейства","Caption",235,true);
            var areaCheck=new FrameworkElementFactory(typeof(CheckBox));
            areaCheck.SetBinding(CheckBox.IsCheckedProperty,new Binding("AreaFromParameter"){Mode=BindingMode.TwoWay,UpdateSourceTrigger=UpdateSourceTrigger.PropertyChanged});
            areaCheck.SetValue(FrameworkElement.HorizontalAlignmentProperty,HorizontalAlignment.Center);
            areaCheck.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);
            areaCheck.SetValue(FrameworkElement.MarginProperty,new Thickness(0));
            areaCheck.SetValue(Control.PaddingProperty,new Thickness(0));
            areaCheck.SetValue(FrameworkElement.ToolTipProperty,"Включено - готовая площадь из параметра экземпляра. Выключено - площадь горизонтальной проекции геометрии.");
            families.Columns.Add(new DataGridTemplateColumn{Header="Из параметра",Width=120,CellTemplate=new DataTemplate{VisualTree=areaCheck}});
            TextColumn(families,"Параметр площади","AreaParameter",150,false,tip:"Имя параметра экземпляра с типом данных Площадь. Применяется при включённом Из параметра.\nПример: АР_Площадь");
            families.ToolTip="Имя параметра площади задаётся отдельно для каждого типа. Нужен параметр экземпляра с типом данных Площадь.\nПример: Площадь";familyContent.Children.Add(families);
            familyContent.Children.Add(new TextBlock{Text="Геометрия даёт контур для расчёта и визуального контроля. Параметр площади даёт только числовую площадь; проверки, которым нужен контур, потребуют геометрию.",TextWrapping=TextWrapping.Wrap,Foreground=new SolidColorBrush(Color.FromRgb(92,111,131)),Margin=new Thickness(0,10,0,0)});
            var footer=new StackPanel{Margin=new Thickness(0,8,0,0)};Grid.SetRow(footer,2);root.Children.Add(footer);
            var status=new TextBlock{TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,8),Text=readError==null?"":"Общий словарь недоступен: "+readError};
            var statusStyle=new Style(typeof(TextBlock));
            var emptyStatus=new Trigger{Property=TextBlock.TextProperty,Value=""};emptyStatus.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Collapsed));statusStyle.Triggers.Add(emptyStatus);
            status.Style=statusStyle;footer.Children.Add(status);
            var buttons=new DockPanel();footer.Children.Add(buttons);
            var right=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right};DockPanel.SetDock(right,Dock.Right);buttons.Children.Add(right);
            var cancel=new Button{Content="Закрыть",IsCancel=true,ToolTip="Закрывает окно без применения настроек. Уже выполненная запись параметров остаётся в модели."};
            var apply=new Button{Content="Применить",ToolTip="Применяет категории к текущему расчёту и обновляет классификацию. Для BIM (8) общий settings.json обновляется автоматически."};right.Children.Add(cancel);right.Children.Add(writeRooms);right.Children.Add(apply);
            string currentRole=null;bool rebinding=false;
            Action save=()=>
            {
                if(currentRole==null)return;
                families.CommitEdit(DataGridEditingUnit.Cell,true);families.CommitEdit(DataGridEditingUnit.Row,true);
                var item=draft.Categories.First(c=>c.Category==roles.First(r=>r.Key==currentRole).Label);
                item.Departments=tags.GetDepartments();item.Names=tags.GetNames();item.DisabledTags=tags.GetDisabledValues();item.ApartmentOnly=apartmentOnly.IsChecked==true;
                draft.Sources=draft.Sources??new List<TEP.CategorySource>();draft.Sources.RemoveAll(c=>c.Role==currentRole);draft.Sources.Add(new TEP.CategorySource{Role=currentRole,Families=familySource.IsChecked==true});
            };
            writeRooms.Click+=(s,e)=>
            {
                try{save();var dialog=BuildApartmentParameterWindow(draft);dialog.Owner=window;dialog.ShowDialog();if(Engine.RequestedDetail!=null)CloseNavigationDialog(window);}
                catch(Exception ex){status.Text=ex.Message;}
            };
            Action load=()=>
            {
                rebinding=true;
                try
                {
                    currentRole=(category.SelectedItem as TEP.Choice)?.Key;if(currentRole==null)return;
                    var item=draft.Categories.First(c=>c.Category==roles.First(r=>r.Key==currentRole).Label);
                    tags.SetValues(item.Departments,item.Names,item.DisabledTags);apartmentOnly.IsChecked=item.ApartmentOnly;
                    bool fromFamilies=draft.Sources?.Any(c=>c.Role==currentRole&&c.Families)==true;familySource.IsChecked=fromFamilies;roomSource.IsChecked=!fromFamilies;
                    families.ItemsSource=(draft.Families??new List<TEP.CategoryFamily>()).Where(f=>f.Role==currentRole).ToList();
                    foreach(var displayed in category.Items.Cast<TEP.Choice>())displayed.Label=roles.First(r=>r.Key==displayed.Key).Label+" | "+(draft.Sources?.Any(c=>c.Role==displayed.Key&&c.Families)==true?"Семейства":"Помещение");
                    category.Items.Refresh();
                }
                finally{rebinding=false;}
            };
            category.SelectionChanged+=(s,e)=>{if(rebinding)return;save();load();};
            add.Click+=(s,e)=>
            {
                try
                {
                    save();var choice=picker.SelectedItem as TEP.FamilyChoice;if(choice==null)throw new InvalidOperationException("Выберите тип семейства.");
                    draft.Families=draft.Families??new List<TEP.CategoryFamily>();
                    if(draft.Families.Any(f=>string.Equals(f.Caption,choice.Caption,StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("Тип уже добавлен в одну из категорий. Сначала удалите прежнее назначение.");
                    draft.Families.Add(new TEP.CategoryFamily{Role=currentRole,Category=choice.Category,Family=choice.Family,Type=choice.Type,AreaParameter="Площадь"});
                    draft.Sources.First(c=>c.Role==currentRole).Families=true;load();
                }
                catch(Exception ex){status.Text=ex.Message;}
            };
            remove.Click+=(s,e)=>{var item=families.SelectedItem as TEP.CategoryFamily;if(item==null)return;save();draft.Families.Remove(item);load();};
            apply.Click+=(s,e)=>
            {
                try
                {
                    save();TEP.Engine.ValidateCategorySettings(draft);
                    if(manager&&readError!=null)throw new InvalidOperationException("Общий словарь недоступен: "+readError+". Откройте настройки после восстановления доступа.");
                    if(manager)Engine.PublishCategoryDictionary(draft,revision);
                    Engine.ApplyCategorySettings(draft);window.DialogResult=true;
                }
                catch(Exception ex){status.Text="Настройки не применены: "+ex.Message;}
            };
            Action refreshLabels=()=>
            {
                if(rebinding)return;save();rebinding=true;
                var chosen=currentRole;
                category.ItemsSource=roles.Select(r=>new TEP.Choice(r.Key,r.Label+" | "+(draft.Sources?.Any(c=>c.Role==r.Key&&c.Families)==true?"Семейства":"Помещение"),r.Description)).ToList();
                category.SelectedValue=chosen;rebinding=false;
            };
            roomSource.Click+=(s,e)=>refreshLabels();familySource.Click+=(s,e)=>refreshLabels();
            category.ItemsSource=roles.Select(r=>new TEP.Choice(r.Key,r.Label+" | "+(draft.Sources?.Any(c=>c.Role==r.Key&&c.Families)==true?"Семейства":"Помещение"),r.Description)).ToList();
            category.SelectedValue=selectedDepartmentRole=="unknown"?roles.First().Key:selectedDepartmentRole;
            load();window.Content=root;return window;
        }
        private void ShowRoomParameterWrite(Window owner,List<TEP.RoomCategoryWrite> selected)
        {
            const string parameter=TEP.Engine.DefaultRoomOutputParameter;
            var rows=Engine.PreviewRoomCategoryWrite(selected);
            var window=new Window{Title="KPLN | Запись назначения: "+parameter,Owner=owner,Width=1150,Height=520,MinWidth=850,MinHeight=380,WindowStartupLocation=WindowStartupLocation.CenterOwner,FontSize=14};
            window.Resources.MergedDictionaries.Add(Resources);
            var root=new DockPanel{Margin=new Thickness(16)};
            var header=new TextBlock{Text="Параметр: "+parameter+". Значение - название категории каждого помещения. Доступно для записи: "+rows.Count(r=>r.Writable)+" из "+rows.Count+".",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,12)};DockPanel.SetDock(header,Dock.Top);root.Children.Add(header);
            var footer=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,12,0,0)};DockPanel.SetDock(footer,Dock.Bottom);root.Children.Add(footer);
            var cancel=new Button{Content="Отмена",IsCancel=true};var write=new Button{Content="Записать в "+rows.Count(r=>r.Writable)+" помещений",IsEnabled=rows.Any(r=>r.Writable),ToolTip="Создаёт отсутствующий параметр и записывает показанное значение в доступные помещения основной модели. Операция доступна для отмены средствами Revit."};footer.Children.Add(cancel);footer.Children.Add(write);
            var grid=new DataGrid{IsReadOnly=true,ItemsSource=rows};TextColumn(grid,"Источник","Source",170,true);TextColumn(grid,"ID","Element",85,true);TextColumn(grid,"Имя","Name",160,true);TextColumn(grid,"Было","Before",130,true);TextColumn(grid,"Будет","After",130,true);TextColumn(grid,"Состояние","Status",280,true);root.Children.Add(grid);
            footer.Children.Insert(0,CreateRoomNavigationButton(grid,window));
            write.Click+=(s,e)=>{try{Engine.WriteRoomCategories(selected);window.DialogResult=true;}catch(Exception ex){MessageBox.Show(window,ex.Message,"KPLN | Запись не выполнена",MessageBoxButton.OK,MessageBoxImage.Error);}};
            window.Content=root;window.ShowDialog();
        }
    }
}
