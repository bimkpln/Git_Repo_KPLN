using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private void AssignedDepartments_Click(object sender,RoutedEventArgs e)
        {
            var window=BuildAssignedDepartmentsWindow();window.Owner=this;window.ShowDialog();if(Engine.RequestedDetail!=null)CloseNavigationDialog(this);else RefreshDepartmentCards();
        }
        private Window BuildAssignedDepartmentsWindow()
        {
            var window=new Window{Title="KPLN | Распределённые назначения",Width=1250,Height=800,MinWidth=960,MinHeight=620,WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=Background,FontSize=14,FontFamily=FontFamily};
            window.Resources.MergedDictionaries.Add(Resources);
            var root=new Grid{Margin=new Thickness(18)};
            root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});root.RowDefinitions.Add(new RowDefinition());root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});
            root.Children.Add(new TextBlock{Text="Выберите категорию. Назначения можно перенести кнопкой или перетащить на другую категорию.",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,14)});
            var body=new Grid();body.ColumnDefinitions.Add(new ColumnDefinition{Width=new GridLength(5,GridUnitType.Star)});body.ColumnDefinitions.Add(new ColumnDefinition{Width=new GridLength(18)});body.ColumnDefinitions.Add(new ColumnDefinition{Width=new GridLength(6,GridUnitType.Star)});Grid.SetRow(body,1);root.Children.Add(body);
            var cards=new ListBox{Tag="distribution-categories",HorizontalContentAlignment=HorizontalAlignment.Stretch,Padding=new Thickness(4)};
            var plainItem=new Style(typeof(ListBoxItem));
            plainItem.Setters.Add(new Setter(Control.PaddingProperty,new Thickness(10,9,10,9)));
            plainItem.Setters.Add(new Setter(Control.BackgroundProperty,Brushes.Transparent));
            var frame=new FrameworkElementFactory(typeof(Border));frame.SetValue(Border.PaddingProperty,new Thickness(10,9,10,9));
            var presenter=new FrameworkElementFactory(typeof(ContentPresenter));frame.AppendChild(presenter);
            plainItem.Setters.Add(new Setter(Control.TemplateProperty,new ControlTemplate(typeof(ListBoxItem)){VisualTree=frame}));cards.ItemContainerStyle=plainItem;body.Children.Add(cards);
            var detail=new Grid();Grid.SetColumn(detail,2);body.Children.Add(detail);
            detail.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});detail.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});detail.RowDefinitions.Add(new RowDefinition{Height=new GridLength(170)});detail.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});detail.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});detail.RowDefinitions.Add(new RowDefinition());
            var title=new TextBlock{FontSize=17,FontWeight=FontWeights.SemiBold,TextWrapping=TextWrapping.Wrap};detail.Children.Add(title);
            var description=new TextBlock{TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,8,0,10)};Grid.SetRow(description,1);detail.Children.Add(description);
            var values=new ListBox{DisplayMemberPath="Caption",SelectionMode=SelectionMode.Extended,ToolTip="Выберите одно или несколько назначений для переноса. Ctrl и Shift позволяют выбрать несколько строк."};Grid.SetRow(values,2);detail.Children.Add(values);
            var movePanel=new DockPanel{Margin=new Thickness(0,10,0,0)};Grid.SetRow(movePanel,3);detail.Children.Add(movePanel);
            var move=new Button{Content="Перенести",Margin=new Thickness(10,0,0,0),IsEnabled=false,ToolTip="Переносит выбранные назначения в указанную категорию. Нераспределённые снимает назначение категории."};DockPanel.SetDock(move,Dock.Right);movePanel.Children.Add(move);
            var target=new ComboBox{ItemsSource=CategoryTargetChoices,DisplayMemberPath="Label",SelectedValuePath="Key",SelectedIndex=0,ToolTip="Категория назначения для выбранных строк. Для возврата выберите Нераспределённые."};movePanel.Children.Add(target);
            var count=new TextBlock{Margin=new Thickness(0,12,0,8)};Grid.SetRow(count,4);detail.Children.Add(count);
            var table=new DataGrid{IsReadOnly=true};TextColumn(table,"Источник","Source",160,true);TextColumn(table,"ID","Element",80,true);TextColumn(table,"Имя","Name",150,true);TextColumn(table,"Этаж","Level",100,true);Grid.SetRow(table,5);detail.Children.Add(table);
            var close=new Button{Content="Закрыть",IsCancel=true,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,14,0,0),ToolTip="Закрывает окно. Выполненные переносы уже применены к текущему расчёту."};var footer=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,14,0,0)};
            close.Margin=new Thickness(0);footer.Children.Add(CreateRoomNavigationButton(table,window));footer.Children.Add(close);Grid.SetRow(footer,2);root.Children.Add(footer);
            var groups=departmentGroups??new List<TEP.DepartmentGroup>();
            string current="unknown";bool updating=false,distributionChanged=false;
            window.Closed+=(s,e)=>{if(distributionChanged)AutoSaveCategoryAssignments();};
            Action showRooms=()=>
            {
                var selected=values.SelectedItems.Cast<TEP.DepartmentGroup>().ToList();
                var visible=selected.Count>0?selected:values.Items.Cast<TEP.DepartmentGroup>().ToList();
                var roomRows=visible.SelectMany(g=>g.Rooms).ToList();
                if(selected.Count==0&&Engine.CategoryUsesFamilies(current))roomRows.AddRange(Engine.CategoryFamilyInstances(current));
                table.ItemsSource=roomRows;count.Text="Объектов: "+roomRows.Count;move.IsEnabled=selected.Count>0;
            };
            Action refresh=null;
            Action<List<TEP.DepartmentGroup>,string> transfer=(items,role)=>
            {
                if(items.Count==0||role==null)return;
                foreach(var group in items)Engine.AssignDepartment(group.Value,role);
                Engine.SetCategorySource(role,false);checkedConfiguration=null;distributionChanged=true;current=role;refresh();RefreshDepartmentCards();
            };
            refresh=()=>
            {
                updating=true;cards.Items.Clear();
                var ordered=DepartmentRoles.Select(role=>new {Role=role,Count=groups.Where(g=>Engine.DepartmentRole(g.Value)==role.Key).Sum(g=>g.Count)+(Engine.CategoryUsesFamilies(role.Key)?Engine.CategoryFamilyInstances(role.Key).Count:0)})
                    .OrderBy(x=>x.Role.Key=="unknown"?0:1).ThenByDescending(x=>x.Count).ThenBy(x=>x.Role.Label,StringComparer.CurrentCultureIgnoreCase);
                foreach(var entry in ordered)
                {
                    var role=entry.Role;
                    var text=new TextBlock{Text=role.Label+" ("+entry.Count+")",TextWrapping=TextWrapping.Wrap};
                    if(role.Key=="unknown"&&entry.Count>0)text.Foreground=ErrorTextBrush;
                    else if(role.Key==current)text.Foreground=new SolidColorBrush(Color.FromRgb(36,92,145));
                    if(role.Key==current)text.FontWeight=FontWeights.SemiBold;
                    var item=new ListBoxItem{Content=text,AllowDrop=true,Tag=role.Key,ToolTip=role.Description};
                    item.Drop+=(s,e)=>{var items=e.Data.GetData(typeof(List<TEP.DepartmentGroup>)) as List<TEP.DepartmentGroup>;if(items==null)return;transfer(items,(string)item.Tag);e.Handled=true;};
                    cards.Items.Add(item);if(role.Key==current)cards.SelectedItem=item;
                }
                updating=false;
                var destination=target.SelectedValue;target.ItemsSource=CategoryTargetChoices;target.SelectedValue=destination??"unknown";
                var choice=DepartmentRoles.First(r=>r.Key==current);title.Text=choice.Label;
                description.Text=Engine.CategoryUsesFamilies(current)?"Расчёт этой категории задан через семейства. Помещения ниже сохраняют распределение, но не суммируются. Перенос назначений в категорию переключает её на помещения.":choice.Description;
                values.ItemsSource=groups.Where(g=>Engine.DepartmentRole(g.Value)==current).ToList();showRooms();
            };
            cards.SelectionChanged+=(s,e)=>{if(updating)return;var item=cards.SelectedItem as ListBoxItem;if(item==null)return;current=(string)item.Tag;refresh();};
            values.SelectionChanged+=(s,e)=>showRooms();
            Point? dragStart=null;values.PreviewMouseLeftButtonDown+=(s,e)=>dragStart=e.GetPosition(values);
            values.PreviewMouseMove+=(s,e)=>
            {
                if(e.LeftButton!=MouseButtonState.Pressed||!dragStart.HasValue||values.SelectedItems.Count==0)return;
                var delta=e.GetPosition(values)-dragStart.Value;
                if(Math.Abs(delta.X)<SystemParameters.MinimumHorizontalDragDistance&&Math.Abs(delta.Y)<SystemParameters.MinimumVerticalDragDistance)return;
                dragStart=null;var selected=values.SelectedItems.Cast<TEP.DepartmentGroup>().ToList();DragDrop.DoDragDrop(values,new DataObject(typeof(List<TEP.DepartmentGroup>),selected),DragDropEffects.Move);
            };
            move.Click+=(s,e)=>transfer(values.SelectedItems.Cast<TEP.DepartmentGroup>().ToList(),target.SelectedValue as string);
            refresh();window.Content=root;return window;
        }
        private sealed class CategoryTagEditor : StackPanel
        {
            private static readonly Style toggleStyle=CreateToggleStyle();
            private static Style CreateToggleStyle()
            {
                var style=new Style(typeof(ToggleButton));
                var border=new FrameworkElementFactory(typeof(Border));border.SetValue(Border.CornerRadiusProperty,new CornerRadius(3));border.SetValue(Border.BorderThicknessProperty,new Thickness(1));border.SetValue(Border.BorderBrushProperty,new SolidColorBrush(Color.FromRgb(157,179,201)));
                border.SetBinding(Border.BackgroundProperty,new Binding("Background"){RelativeSource=new RelativeSource(RelativeSourceMode.TemplatedParent)});
                var content=new FrameworkElementFactory(typeof(ContentPresenter));content.SetValue(FrameworkElement.HorizontalAlignmentProperty,HorizontalAlignment.Center);content.SetValue(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center);border.AppendChild(content);
                style.Setters.Add(new Setter(Control.TemplateProperty,new ControlTemplate(typeof(ToggleButton)){VisualTree=border}));return style;
            }
            private sealed class TagValue {internal string Value;internal bool Department;internal bool Name;}
            private readonly List<TagValue> values=new List<TagValue>();
            private readonly WrapPanel tags=new WrapPanel();
            private readonly DockPanel entry=new DockPanel{Margin=new Thickness(0,0,6,6)};
            private readonly TextBox input;
            internal CategoryTagEditor()
            {
                Children.Add(new ScrollViewer{Content=tags,MaxHeight=240,Padding=new Thickness(0,4,6,4),VerticalScrollBarVisibility=ScrollBarVisibility.Auto,HorizontalScrollBarVisibility=ScrollBarVisibility.Disabled});
                input=new TextBox{Width=165,MinHeight=32,Padding=new Thickness(7,3,7,3),ToolTip="Введите значение. Новый тег сравнивается и с Назначением, и с Именем комнаты.\nПример: Санузел"};
                var add=new Button{Content="+",Width=30,MinWidth=30,MinHeight=32,Padding=new Thickness(0),Margin=new Thickness(4,0,0,0),ToolTip="Добавить тег; Н и К включены по умолчанию."};DockPanel.SetDock(add,Dock.Right);entry.Children.Add(add);entry.Children.Add(input);
                add.Click+=(s,e)=>Commit();input.KeyDown+=(s,e)=>{if(e.Key==Key.Enter){Commit();e.Handled=true;}};RefreshTags();
            }
            internal List<string> GetDepartments(){Commit();return values.Where(v=>v.Department).Select(v=>v.Value).ToList();}
            internal List<string> GetNames(){Commit();return values.Where(v=>v.Name).Select(v=>v.Value).ToList();}
            internal List<string> GetDisabledValues(){Commit();return values.Where(v=>!v.Department&&!v.Name).Select(v=>v.Value).ToList();}
            internal void SetValues(IEnumerable<string> departments,IEnumerable<string> names,IEnumerable<string> disabled)
            {
                values.Clear();
                foreach(var source in new[]{Tuple.Create(departments,true),Tuple.Create(names,false)})
                    foreach(string text in source.Item1??Enumerable.Empty<string>())
                    {
                        if(string.IsNullOrWhiteSpace(text))continue;string clean=text.Trim();
                        var tag=values.FirstOrDefault(v=>v.Value.Equals(clean,StringComparison.OrdinalIgnoreCase));
                        if(tag==null){tag=new TagValue{Value=clean};values.Add(tag);}
                        if(source.Item2)tag.Department=true;else tag.Name=true;
                    }
                foreach(string text in disabled??Enumerable.Empty<string>())if(!string.IsNullOrWhiteSpace(text)&&!values.Any(v=>v.Value.Equals(text.Trim(),StringComparison.OrdinalIgnoreCase)))values.Add(new TagValue{Value=text.Trim()});
                input.Clear();RefreshTags();
            }
            private void Commit()
            {
                string text=input.Text.Trim();if(text.Length==0)return;
                if(!values.Any(v=>v.Value.Equals(text,StringComparison.OrdinalIgnoreCase)))values.Add(new TagValue{Value=text,Department=true,Name=true});
                input.Clear();RefreshTags();input.Focus();
            }
            private void RefreshTags()
            {
                tags.Children.Clear();
                foreach(var value in values)
                {
                    var content=new StackPanel{Orientation=Orientation.Horizontal};
                    content.Children.Add(new TextBlock{Text=value.Value,MaxWidth=230,TextTrimming=TextTrimming.CharacterEllipsis,ToolTip=value.Value,VerticalAlignment=VerticalAlignment.Center,Margin=new Thickness(0,0,6,0)});
                    foreach(bool department in new[]{true,false})
                    {
                        var toggle=new ToggleButton{Style=toggleStyle,Content=department?"Н":"К",Tag=(department?"Н|":"К|")+value.Value,IsChecked=department?value.Department:value.Name,Width=27,MinHeight=26,Padding=new Thickness(0),Margin=new Thickness(0,0,3,0),ToolTip=department?"Н: сравнивать с параметром Назначение":"К: сравнивать с Именем комнаты"};
                        Action color=()=>{toggle.Foreground=toggle.IsChecked==true?Brushes.White:new SolidColorBrush(Color.FromRgb(90,103,119));toggle.Background=toggle.IsChecked==true?new SolidColorBrush(Color.FromRgb(36,92,145)):Brushes.Transparent;};
                        toggle.Click+=(s,e)=>{if(department)value.Department=toggle.IsChecked==true;else value.Name=toggle.IsChecked==true;color();};color();content.Children.Add(toggle);
                    }
                    var remove=new Button{Content="×",MinWidth=24,MinHeight=26,Padding=new Thickness(3,0,3,0),Margin=new Thickness(2,0,0,0),ToolTip="Удалить тег «"+value.Value+"» из обоих правил",Tag=value.Value};content.Children.Add(remove);
                    var border=new Border{Child=content,Background=new SolidColorBrush(Color.FromRgb(232,238,245)),CornerRadius=new CornerRadius(8),Padding=new Thickness(8,5,6,5),Margin=new Thickness(0,0,6,6)};tags.Children.Add(border);
                    remove.Click+=(s,e)=>{values.Remove(value);RefreshTags();};
                }
                tags.Children.Add(entry);
            }
        }
    }
}
