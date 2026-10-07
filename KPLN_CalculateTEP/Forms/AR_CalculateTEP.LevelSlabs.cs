using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private DataGridTemplateColumn topSlabColumn;
        private System.Collections.ObjectModel.ObservableCollection<TEP.LevelSetting> observedLevels;
        private readonly System.Collections.Generic.List<TEP.LevelSetting> observedSlabRows=new System.Collections.Generic.List<TEP.LevelSetting>();
        private void StopObservingSlabRows()
        {
            if(observedLevels!=null)observedLevels.CollectionChanged-=SlabLevels_Changed;
            foreach(var row in observedSlabRows)row.PropertyChanged-=SlabRequirement_Changed;
            observedSlabRows.Clear();observedLevels=null;
        }
        private void ObserveSlabRows()
        {
            foreach(var row in observedSlabRows)row.PropertyChanged-=SlabRequirement_Changed;
            observedSlabRows.Clear();
            foreach(var row in Config.Levels){row.PropertyChanged+=SlabRequirement_Changed;observedSlabRows.Add(row);}
            UpdateSlabColumnVisibility();
        }
        private void SlabLevels_Changed(object sender,System.Collections.Specialized.NotifyCollectionChangedEventArgs e){ObserveSlabRows();}
        private void SlabRequirement_Changed(object sender,System.ComponentModel.PropertyChangedEventArgs e)
        {if(e.PropertyName.EndsWith("Required"))UpdateSlabColumnVisibility();}
        private void UpdateSlabColumnVisibility()
        {if(topSlabColumn!=null)topSlabColumn.Visibility=Config.Levels.Any(l=>l.TopSlabRequired)?Visibility.Visible:Visibility.Collapsed;UpdateLevelGeometryColumns();}
        private void TopSlabColumn()
        {
            var button=new FrameworkElementFactory(typeof(Button));
            button.SetBinding(ContentControl.ContentProperty,new Binding("TopSlabSummary"));
            var style=new Style(typeof(Button),TryFindResource(typeof(Button)) as Style);
            style.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Collapsed));
            var required=new DataTrigger{Binding=new Binding("TopSlabRequired"),Value=true};
            required.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Visible));style.Triggers.Add(required);
            button.SetValue(FrameworkElement.StyleProperty,style);
            button.SetValue(FrameworkElement.ToolTipProperty,"Верх перекрытия над цоколем определяется по геометрии. Нажмите, чтобы увидеть отметки, источники и ID, выбрать вариант или задать абсолютную отметку вручную.");
            button.SetValue(Control.PaddingProperty,new Thickness(6,2,6,2));button.SetValue(FrameworkElement.MarginProperty,new Thickness(3,2,3,2));
            button.AddHandler(Button.ClickEvent,new RoutedEventHandler(TopSlab_Click));
            topSlabColumn=new DataGridTemplateColumn{Header="Верх перекрытия, м",Width=161,CellTemplate=new DataTemplate{VisualTree=button}};
            LevelsGrid.Columns.Add(topSlabColumn);
            if(observedLevels!=null)observedLevels.CollectionChanged-=SlabLevels_Changed;
            observedLevels=Config.Levels;observedLevels.CollectionChanged+=SlabLevels_Changed;ObserveSlabRows();
        }
        private void LevelKind_Changed(object sender,SelectionChangedEventArgs args)
        {
            if(busy||!IsLoaded||!Engine.IsInitialized)return;
            var combo=sender as ComboBox;
            if(combo==null||(!combo.IsKeyboardFocusWithin&&!combo.IsDropDownOpen))return;
            var row=(sender as FrameworkElement)?.DataContext as TEP.LevelSetting;if(row==null)return;
            Dispatcher.BeginInvoke(DispatcherPriority.Background,new Action(()=>
            {
                if(busy)return;
                Try(()=>{SetBusy(true);try{Engine.RefreshTopSlabs(ReportProgress);Engine.PreviewLevelGeometry(ReportProgress);row.TopSlabSummary=row.TopSlabSummary;checkedConfiguration=null;}
                    finally{SetBusy(false);}});
            }));
        }
        private void TopSlab_Click(object sender,RoutedEventArgs args)
        {
            var row=(sender as FrameworkElement)?.DataContext as TEP.LevelSetting;if(row==null)return;
            Try(()=>
            {
                if(row.Kind!="basement"||row.Above!="auto")
                {Status.Text="Верх перекрытия нужен для цоколя с автоматическим определением наземности.";return;}
                SetBusy(true);try{Engine.RefreshTopSlabs(ReportProgress);}finally{SetBusy(false);}
                ShowTopSlabChoice(row);
            });
        }
        private void ResolveAmbiguousTopSlabs()
        {
            foreach(var row in Config.Levels.Where(l=>l.Include&&l.Kind=="basement"&&l.Above=="auto"&&l.TopSlabMode!="manual"))
            {
                var review=Engine.TopSlabReviewFor(row);
                if(review.Selected==null&&review.Options.Any(o=>o.CoversRooms))ShowTopSlabChoice(row);
            }
        }
        private void ShowTopSlabChoice(TEP.LevelSetting row)
        {
            // Manual settings are retained until a user actually confirms another mode.
            if(row.TopSlabMode=="manual")
            {
                var old=row.TopSlabMode;var value=row.TopSlab;row.TopSlabMode="auto";
                SetBusy(true);try{Engine.RefreshTopSlabs(ReportProgress);}
                finally{row.TopSlabMode=old;row.TopSlab=value;row.TopSlabSummary="Вручную: "+value+" м";SetBusy(false);}
            }
            var review=Engine.TopSlabReviewFor(row);
            var dialog=CreateTopSlabDialog(row,review);dialog.Owner=this;
            bool accepted=dialog.ShowDialog()==true;
            if(accepted){checkedConfiguration=null;SetBusy(true);try{Engine.RefreshTopSlabs(ReportProgress);}finally{SetBusy(false);}}
            else if(row.TopSlabMode=="manual")row.TopSlabSummary="Вручную: "+row.TopSlab+" м";
        }
        private Window CreateTopSlabDialog(TEP.LevelSetting row,TEP.TopSlabReview review)
        {
            var dialog=new Window{Title="KPLN | Перекрытие над цоколем: "+row.Name,Width=960,Height=570,MinWidth=700,MinHeight=460,
                WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=Background,FontFamily=FontFamily,FontSize=14};
            var root=new DockPanel{Margin=new Thickness(16)};dialog.Content=root;
            var introduction=new TextBlock{Text="Абсолютные отметки верха перекрытий в координатах основной модели. Вариант объединяет перекрытия с одной отметкой. Выбрать можно вариант, который перекрывает все размещённые помещения этого этажа. Наклонный или разноуровневый верх автоматически одним числом не заменяется.",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,12)};
            DockPanel.SetDock(introduction,Dock.Top);root.Children.Add(introduction);
            var bottom=new StackPanel{Margin=new Thickness(0,12,0,0)};DockPanel.SetDock(bottom,Dock.Bottom);root.Children.Add(bottom);
            var manualRow=new StackPanel{Orientation=Orientation.Horizontal};bottom.Children.Add(manualRow);
            manualRow.Children.Add(new TextBlock{Text="Ручная абсолютная отметка, м",VerticalAlignment=VerticalAlignment.Center,Margin=new Thickness(0,0,12,0)});
            var manual=new TextBox{Text=row.TopSlab??"",Width=130,Padding=new Thickness(5),ToolTip="Используется вместо геометрии только после нажатия «Применить вручную». Влияет на определение наземности цоколя.\nПример: 158,20"};manualRow.Children.Add(manual);
            var actions=new WrapPanel{Margin=new Thickness(0,10,0,0)};bottom.Children.Add(actions);
            var list=new ListBox{ItemsSource=review.Options,DisplayMemberPath="Display",HorizontalContentAlignment=HorizontalAlignment.Stretch};
            ScrollViewer.SetHorizontalScrollBarVisibility(list,ScrollBarVisibility.Disabled);
            var textStyle=new Style(typeof(TextBlock));textStyle.Setters.Add(new Setter(TextBlock.TextWrappingProperty,TextWrapping.Wrap));list.Resources.Add(typeof(TextBlock),textStyle);
            list.SelectedItem=review.Options.FirstOrDefault(o=>o.Key==row.TopSlabChoice)??review.Selected;
            Action<string,Action> add=(caption,action)=>
            {var button=new Button{Content=caption,Margin=new Thickness(0,0,8,0),Padding=new Thickness(12,6,12,6),ToolTip=
                caption=="Выбрать вариант"?"Использовать отметку выбранных перекрытий. При следующем расчёте их геометрия проверяется заново.":
                caption=="Автоматически"?"Отменить ручное назначение. Единственный пригодный вариант подставится автоматически; неоднозначность останется замечанием.":
                caption=="Применить вручную"?"Использовать введённую абсолютную отметку вместо геометрии. В отчёте будет предупреждение о ручном значении.":"Закрыть окно без изменения выбранного способа."};button.Click+=(s,e)=>action();actions.Children.Add(button);};
            add("Выбрать вариант",()=>
            {
                var option=list.SelectedItem as TEP.TopSlabOption;
                if(option==null||!option.CoversRooms){MessageBox.Show(dialog,"Выберите вариант, который перекрывает все помещения. Для частичного покрытия общая отметка не определена.","KPLN | Верх перекрытия");return;}
                row.TopSlabMode="auto";row.TopSlabChoice=option.Key;row.TopSlab=null;dialog.DialogResult=true;
            });
            add("Автоматически",()=>{row.TopSlabMode="auto";row.TopSlabChoice=null;row.TopSlab=null;dialog.DialogResult=true;});
            add("Применить вручную",()=>
            {
                double value;if(!TEP.Engine.TryNumber(manual.Text,out value)){MessageBox.Show(dialog,"Введите конечное число в метрах.","KPLN | Верх перекрытия");return;}
                row.TopSlabMode="manual";row.TopSlab=manual.Text;row.TopSlabChoice=null;dialog.DialogResult=true;
            });
            add("Закрыть",()=>dialog.Close());
            var warning=new TextBlock{Text=review.Error??(review.Problems.Count>0?string.Join("\n",review.Problems):""),TextWrapping=TextWrapping.Wrap,Foreground=ErrorTextBrush,Margin=new Thickness(0,0,0,10)};
            DockPanel.SetDock(warning,Dock.Top);root.Children.Add(new ScrollViewer{Content=warning,MaxHeight=140,VerticalScrollBarVisibility=ScrollBarVisibility.Auto});
            // The warning belongs above the list, which fills the remaining space.
            DockPanel.SetDock(root.Children[root.Children.Count-1],Dock.Top);root.Children.Add(list);
            return dialog;
        }
    }
}
