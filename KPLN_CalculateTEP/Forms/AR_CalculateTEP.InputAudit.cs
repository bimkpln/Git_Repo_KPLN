using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private static void CloseNavigationDialog(Window dialog)
        {
            TEP.Engine.TraceNavigation("Closing dialog: "+dialog.Title);
            dialog.Close();
            TEP.Engine.TraceNavigation("Closed dialog: "+dialog.Title);
        }
        private static TEP.InputAuditRow RoomNavigationTarget(object selected)
        {
            var room=selected as TEP.DepartmentRoom ?? (selected as TEP.RoomCategoryWrite)?.Room;
            if(room!=null)return new TEP.InputAuditRow{SourceKey=room.SourceKey,Source=room.Source,ElementId=room.Element};
            var write=selected as TEP.RoomParameterWrite;
            return write==null?null:new TEP.InputAuditRow{SourceKey=write.SourceKey,Source=write.Source,ElementId=write.Element};
        }
        private Button CreateRoomNavigationButton(DataGrid table,Window owner)
        {
            var button=new Button{Content="Перейти к помещению",Tag="navigate-room",IsEnabled=false,
                ToolTip="Выделяет выбранное помещение и приближает его в Revit. Окна плагина закроются; применённые настройки останутся в памяти. Для связи учитывается её положение в модели."};
            table.SelectionChanged+=(s,e)=>{long id;button.IsEnabled=table.SelectedItems.Count==1&&long.TryParse(RoomNavigationTarget(table.SelectedItem)?.ElementId,out id);};
            button.Click+=(s,e)=>
            {
                try
                {
                    if(table.SelectedItems.Count!=1)return;
                    Engine.RequestAuditNavigation(RoomNavigationTarget(table.SelectedItem));
                    CloseNavigationDialog(owner);
                }
                catch(Exception ex){MessageBox.Show(owner,ex.Message,"KPLN | Переход к помещению",MessageBoxButton.OK,MessageBoxImage.Warning);}
            };
            return button;
        }
        private void DepartmentValues_MouseMove(object sender,MouseEventArgs e)
        {
            if(e.LeftButton!=MouseButtonState.Pressed||DepartmentValues.SelectedItems.Count==0)return;
            var values=DepartmentValues.SelectedItems.Cast<TEP.DepartmentGroup>().ToList();
            DragDrop.DoDragDrop(DepartmentValues,new DataObject(typeof(List<TEP.DepartmentGroup>),values),DragDropEffects.Move);
        }
        private bool ShowInputAudit(TEP.InputAudit audit)
        {
            var window=BuildInputAuditWindow(audit);
            bool accepted=window.ShowDialog()==true;
            if(Engine.RequestedDetail!=null)CloseNavigationDialog(this);
            return accepted;
        }
        private Window BuildInputAuditWindow(TEP.InputAudit audit)
        {
            var window=new Window{Owner=IsLoaded?this:null,Title="KPLN | Проверка исходных данных",Width=Math.Min(1500,SystemParameters.WorkArea.Width-40),Height=Math.Min(820,SystemParameters.WorkArea.Height-40),
                MinWidth=800,MinHeight=500,WindowStartupLocation=WindowStartupLocation.CenterOwner,Resources=Resources,Background=Background,FontFamily=FontFamily,FontSize=FontSize,Language=Language};
            var root=new Grid{Margin=new Thickness(18),Background=Background};
            root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});
            root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});
            root.RowDefinitions.Add(new RowDefinition{Height=new GridLength(1,GridUnitType.Star)});
            root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});
            root.RowDefinitions.Add(new RowDefinition{Height=GridLength.Auto});
            bool navigationRequested=false;
            Action<TEP.InputAuditRow> navigate=row=>
            {
                if(navigationRequested)return;
                try
                {
                    Engine.RequestAuditNavigation(row);
                    navigationRequested=true;
                    CloseNavigationDialog(window);
                }
                catch(Exception ex){MessageBox.Show(window,ex.Message,"KPLN | ТЭП: переход к объекту",MessageBoxButton.OK,MessageBoxImage.Warning);}
            };
            var table=CreateInputAuditTable(audit,navigate);
            Grid.SetRow(table,2);root.Children.Add(table);
            var single=new CheckBox{Content="Считать выбранные источники одним корпусом «"+(audit.Building?.BuildingName??"Единый корпус")+"»",
                IsEnabled=audit.Building?.CanAssumeSingle==true,IsChecked=false,Margin=new Thickness(0,12,0,0),
                Visibility=audit.Building!=null&&audit.Building.Missing+audit.Building.Empty>0?Visibility.Visible:Visibility.Collapsed,
                ToolTip="Только явное подтверждение разрешает заменить отсутствующий ПОМ_Корпус единым корпусом внутри расчёта. Параметры RVT не меняются. При разных корпусах или недоступных источниках допущение запрещено."};
            Grid.SetRow(single,3);root.Children.Add(single);
            var buttons=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,12,0,0)};
            var back=new Button{Content="Вернуться к настройкам",IsCancel=true,MinWidth=180};
            var calculate=new Button{Content="Построить и рассчитать",MinWidth=190,Margin=new Thickness(12,0,0,0),ToolTip="Продолжает расчёт доступных объектов и показателей, даже если таблица содержит ошибки. Пропуски и неполные итоги отражаются в результате."};
            calculate.Click+=(s,e)=>{Engine.AcceptAuditBuilding(audit,single.IsChecked==true);window.DialogResult=true;};
            buttons.Children.Add(back);buttons.Children.Add(calculate);Grid.SetRow(buttons,4);root.Children.Add(buttons);window.Content=root;
            return window;
        }
        private DataGrid CreateInputAuditTable(TEP.InputAudit audit,Action<TEP.InputAuditRow> navigate)
        {
            var table=new DataGrid{IsReadOnly=true,AutoGenerateColumns=false,EnableRowVirtualization=true,EnableColumnVirtualization=false,CanUserAddRows=false,MinWidth=0,MinHeight=0,HorizontalAlignment=HorizontalAlignment.Stretch,VerticalAlignment=VerticalAlignment.Stretch};
            VirtualizingPanel.SetIsVirtualizingWhenGrouping(table,true);
            foreach(var column in new[]{Tuple.Create("Источник","Source",240.0),Tuple.Create("Имя","Name",240.0),Tuple.Create("Параметр / проверка","Parameter",250.0),Tuple.Create("Инфо","Info",350.0)})
            {
                var style=new Style(typeof(TextBlock));style.Setters.Add(new Setter(TextBlock.TextWrappingProperty,TextWrapping.Wrap));
                var problem=new DataTrigger{Binding=new Binding("Problem"),Value=true};problem.Setters.Add(new Setter(TextBlock.ForegroundProperty,ErrorTextBrush));style.Triggers.Add(problem);
                style.Setters.Add(new Setter(FrameworkElement.MarginProperty,new Thickness(6,4,6,4)));
                style.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center));
                if(column.Item2=="Parameter")style.Setters.Add(new Setter(FrameworkElement.ToolTipProperty,new Binding("TechnicalCode")));
                table.Columns.Add(new DataGridTextColumn{Header=column.Item1,Binding=new Binding(column.Item2),SortMemberPath=column.Item2,Width=new DataGridLength(column.Item3),MinWidth=column.Item2=="Info"?300:column.Item2=="Parameter"?220:200,ElementStyle=style});
            }
            var idButton=new FrameworkElementFactory(typeof(Button));idButton.SetBinding(Button.ContentProperty,new Binding("ElementId"));
            idButton.SetValue(Button.ToolTipProperty,"Перейти к элементу в Revit. Окна проверки и плагина закроются, чтобы можно было работать с моделью. Настройки сохраняются в памяти до закрытия Revit.");
            idButton.AddHandler(Button.ClickEvent,new RoutedEventHandler((s,e)=>navigate(((FrameworkElement)s).DataContext as TEP.InputAuditRow)));
            table.Columns.Insert(0,new DataGridTemplateColumn{Header="Элемент ID",CellTemplate=new DataTemplate{VisualTree=idButton},Width=100,MinWidth=100,SortMemberPath="ElementId",CanUserSort=true});
            // Grouped WPF grids can permanently shrink pixel columns if a star column is
            // measured before the table has a viewport. Keep every width in pixels.
            bool widthRefreshPending=false;
            Action fitWidths=()=>
            {
                FitAuditInfoColumn(table);
                if(widthRefreshPending)return;
                widthRefreshPending=true;
                table.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle,new Action(()=>
                {widthRefreshPending=false;InvalidateAuditPanelWidths(table);}));
            };
            table.Loaded+=(s,e)=>fitWidths();
            table.SizeChanged+=(s,e)=>{if(e.WidthChanged)fitWidths();};
            table.MouseDoubleClick+=(s,e)=>{var row=ItemsControl.ContainerFromElement(table,e.OriginalSource as DependencyObject) as DataGridRow;if(row?.Item is TEP.InputAuditRow item)navigate(item);};
            table.RowHeight=double.NaN;table.MinRowHeight=30;
            var groupHeader=new FrameworkElementFactory(typeof(TextBlock));
            var groupTitle=new MultiBinding{StringFormat="{0} ({1})"};groupTitle.Bindings.Add(new Binding("Items[0].ProblemGroupTitle"));groupTitle.Bindings.Add(new Binding("ItemCount"));
            groupHeader.SetBinding(TextBlock.TextProperty,groupTitle);
            groupHeader.SetValue(TextBlock.TextWrappingProperty,TextWrapping.Wrap);groupHeader.SetValue(TextBlock.FontWeightProperty,FontWeights.SemiBold);
            groupHeader.SetValue(FrameworkElement.MarginProperty,new Thickness(6));
            groupHeader.SetBinding(FrameworkElement.MaxWidthProperty,new Binding("ActualWidth"){Source=table,Converter=new AuditHeaderWidthConverter()});
            var expander=new FrameworkElementFactory(typeof(Expander));expander.SetValue(Expander.IsExpandedProperty,false);
            expander.SetValue(Control.HorizontalContentAlignmentProperty,HorizontalAlignment.Stretch);
            expander.SetValue(HeaderedContentControl.HeaderTemplateProperty,new DataTemplate{VisualTree=groupHeader});expander.SetBinding(HeaderedContentControl.HeaderProperty,new Binding());
            expander.AppendChild(new FrameworkElementFactory(typeof(ItemsPresenter)));
            var groupStyle=new Style(typeof(GroupItem));groupStyle.Setters.Add(new Setter(Control.TemplateProperty,new ControlTemplate(typeof(GroupItem)){VisualTree=expander}));
            var groupWidth=new MultiBinding{Converter=new AuditGroupWidthConverter()};
            groupWidth.Bindings.Add(new Binding("ActualWidth"){Source=table});
            foreach(var column in table.Columns)groupWidth.Bindings.Add(new Binding("ActualWidth"){Source=column});
            groupStyle.Setters.Add(new Setter(FrameworkElement.WidthProperty,groupWidth));
            table.GroupStyle.Add(new GroupStyle{ContainerStyle=groupStyle});
            var view=new ListCollectionView(audit.Rows.Where(r=>r.Problem||r.Critical).ToList());table.ItemsSource=view;
            string sortMember="ElementId";bool descending=false;
            Action refresh=()=>
            {
                using(view.DeferRefresh())
                {
                    view.GroupDescriptions.Clear();
                    view.GroupDescriptions.Add(new PropertyGroupDescription("ProblemGroupKey"));
                    view.CustomSort=AuditRowComparer(audit.Rows,sortMember,descending);
                }
            };
            table.Sorting+=(s,e)=>
            {
                e.Handled=true;sortMember=e.Column.SortMemberPath;descending=e.Column.SortDirection==ListSortDirection.Ascending;
                foreach(var column in table.Columns)column.SortDirection=null;
                e.Column.SortDirection=descending?ListSortDirection.Descending:ListSortDirection.Ascending;refresh();
            };
            refresh();
            return table;
        }
        private static void FitAuditInfoColumn(DataGrid table)
        {
            if(table.ActualWidth<=0||table.Columns.Count==0)return;
            var info=table.Columns.Last();
            double occupied=table.Columns.Take(table.Columns.Count-1).Sum(c=>c.Width.Value);
            double width=Math.Max(info.MinWidth,table.ActualWidth-table.BorderThickness.Left-table.BorderThickness.Right-SystemParameters.VerticalScrollBarWidth-occupied);
            if(Math.Abs(info.Width.Value-width)>.5)info.Width=new DataGridLength(width);
        }
        private static void InvalidateAuditPanelWidths(DependencyObject parent)
        {
            // Group virtualization retains the old horizontal extent after a window shrink.
            // Remeasure its panels without refreshing items or losing expanded/selected rows.
            for(int i=0;i<VisualTreeHelper.GetChildrenCount(parent);i++)
            {
                var child=VisualTreeHelper.GetChild(parent,i);
                if(child is VirtualizingStackPanel panel)panel.InvalidateMeasure();
                InvalidateAuditPanelWidths(child);
            }
        }
        private sealed class AuditHeaderWidthConverter:IValueConverter
        {
            public object Convert(object value,Type targetType,object parameter,System.Globalization.CultureInfo culture)
            {return Math.Max(0,(value is double width?width:0)-48);}
            public object ConvertBack(object value,Type targetType,object parameter,System.Globalization.CultureInfo culture)
            {return Binding.DoNothing;}
        }
        private sealed class AuditGroupWidthConverter:IMultiValueConverter
        {
            public object Convert(object[] values,Type targetType,object parameter,System.Globalization.CultureInfo culture)
            {
                double viewport=values[0] is double width?width-SystemParameters.VerticalScrollBarWidth-2:0;
                return Math.Max(Math.Max(0,viewport),values.Skip(1).OfType<double>().Sum());
            }
            public object[] ConvertBack(object value,Type[] targetTypes,object parameter,System.Globalization.CultureInfo culture)
            {throw new NotSupportedException();}
        }
        private static System.Collections.IComparer AuditRowComparer(IEnumerable<TEP.InputAuditRow> rows,string member,bool descending)
        {
            var property=typeof(TEP.InputAuditRow).GetProperty(member);
            Func<TEP.InputAuditRow,string> value=row=>property?.GetValue(row) as string??"";
            Func<string,string,int> compare=(a,b)=>
            {
                long first,second;
                int result=member=="ElementId"&&long.TryParse(a,out first)&&long.TryParse(b,out second)?first.CompareTo(second):StringComparer.CurrentCultureIgnoreCase.Compare(a,b);
                return descending?-result:result;
            };
            return (System.Collections.IComparer)Comparer<TEP.InputAuditRow>.Create((a,b)=>
            {
                int result;
                if(a.ProblemGroupKey!=b.ProblemGroupKey)
                {
                    result=a.SortPriority.CompareTo(b.SortPriority);if(result!=0)return result;
                    return StringComparer.CurrentCulture.Compare(a.ProblemGroupTitle,b.ProblemGroupTitle);
                }
                result=a.SortPriority.CompareTo(b.SortPriority);if(result!=0)return result;
                result=compare(value(a),value(b));if(result!=0)return result;
                result=StringComparer.CurrentCultureIgnoreCase.Compare(a.Parameter,b.Parameter);if(result!=0)return result;
                return StringComparer.Ordinal.Compare(a.ElementGroupKey,b.ElementGroupKey);
            });
        }
    }
}
