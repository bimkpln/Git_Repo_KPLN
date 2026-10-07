using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private DataGridTemplateColumn heightColumn,roofAreaColumn,roofRatioColumn;
        private void NormativeProfile_Changed(object sender,SelectionChangedEventArgs args)
        {
            RefreshRequiredParameters();
            foreach(var row in Config.Levels)row.NormativeProfile=Config.Profile;
            var combo=sender as ComboBox;
            if(!IsLoaded||busy||!Engine.IsInitialized||combo==null||(!combo.IsKeyboardFocusWithin&&!combo.IsDropDownOpen))return;
            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background,new Action(()=>
                Try(()=>{SetBusy(true);try{Engine.PreviewLevelGeometry(ReportProgress);}finally{SetBusy(false);}})));
        }
        private void UpdateLevelGeometryColumns()
        {
            if(heightColumn!=null)heightColumn.Visibility=Config.Levels.Any(l=>l.HeightRequired)?Visibility.Visible:Visibility.Collapsed;
            if(roofAreaColumn!=null)roofAreaColumn.Visibility=Config.Levels.Any(l=>l.RoofAreaRequired)?Visibility.Visible:Visibility.Collapsed;
            if(roofRatioColumn!=null)roofRatioColumn.Visibility=Config.Levels.Any(l=>l.RoofRatioRequired)?Visibility.Visible:Visibility.Collapsed;
        }
        private DataGridTemplateColumn LevelGeometryColumn(string title,string value,string required,double width,string hint)
        {
            var button=new FrameworkElementFactory(typeof(Button));button.SetBinding(ContentControl.ContentProperty,new Binding(value));
            button.SetValue(FrameworkElement.ToolTipProperty,hint+" Нажмите для просмотра исходных данных и уточнения.");
            button.SetValue(Control.PaddingProperty,new Thickness(5,2,5,2));button.SetValue(FrameworkElement.MarginProperty,new Thickness(3,2,3,2));
            var style=new Style(typeof(Button),TryFindResource(typeof(Button)) as Style);style.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Collapsed));
            var trigger=new DataTrigger{Binding=new Binding(required),Value=true};trigger.Setters.Add(new Setter(UIElement.VisibilityProperty,Visibility.Visible));style.Triggers.Add(trigger);
            button.SetValue(FrameworkElement.StyleProperty,style);button.AddHandler(Button.ClickEvent,new RoutedEventHandler(LevelGeometry_Click));
            var column=new DataGridTemplateColumn{Header=title,Width=width,CellTemplate=new DataTemplate{VisualTree=button}};LevelsGrid.Columns.Add(column);return column;
        }
        private void LevelGeometryColumns()
        {
            foreach(var row in Config.Levels)row.NormativeProfile=Config.Profile;
            heightColumn=LevelGeometryColumn("Высота, м","HeightSummary","HeightRequired",130,"Высота по геометрии пола и верхней границы. При перепаде показывается диапазон; одно значение требует уточнения.");
            roofRatioColumn=LevelGeometryColumn("Доля от кровли","RoofRatioSummary","RoofRatioRequired",140,"Площадь технических помещений, делённая на площадь выбранной кровли в плане. Пересечения площадей не суммируются дважды.");
            roofAreaColumn=LevelGeometryColumn("Площадь надстройки, м²","RoofAreaSummary","RoofAreaRequired",170,"Площадь по наружному 2D-контуру надстройки. Область доступна для редактирования в визуальном контроле после расчёта.");
            UpdateLevelGeometryColumns();
        }
        private void LevelGeometry_Click(object sender,RoutedEventArgs args)
        {
            var row=(sender as FrameworkElement)?.DataContext as TEP.LevelSetting;if(row==null)return;
            Try(()=>{SetBusy(true);try{Engine.PreviewLevelGeometry(ReportProgress);}finally{SetBusy(false);}ShowLevelGeometry(row);});
        }
        private void ResolveLevelGeometryChoices()
        {
            foreach(var row in Config.Levels.Where(l=>l.Include))
            {
                var review=Engine.LevelGeometryFor(row);
                if(row.HeightRequired&&row.HeightMode!="manual"&&review.HeightMin.HasValue&&review.HeightMax-review.HeightMin>1e-6||
                    row.RoofRatioRequired&&row.RoofRatioMode!="manual"&&review.Roofs.Count>1&&(row.RoofSources==null||row.RoofSources.Count==0))ShowLevelGeometry(row);
            }
        }
        private void ShowLevelGeometry(TEP.LevelSetting row)
        {
            var dialog=CreateLevelGeometryDialog(row,Engine.LevelGeometryFor(row));dialog.Owner=this;
            if(dialog.ShowDialog()==true)
            {checkedConfiguration=null;SetBusy(true);try{Engine.PreviewLevelGeometry(ReportProgress);}finally{SetBusy(false);}}
        }
        private Window CreateLevelGeometryDialog(TEP.LevelSetting row,TEP.LevelGeometryReview review)
        {
            var dialog=new Window{Title="KPLN | Геометрия этажа: "+row.Name,Width=1040,Height=730,MinWidth=800,MinHeight=580,
                Background=Background,FontFamily=FontFamily,FontSize=14,WindowStartupLocation=WindowStartupLocation.CenterOwner};
            var root=new DockPanel{Margin=new Thickness(16)};dialog.Content=root;
            var actions=new WrapPanel{Margin=new Thickness(0,12,0,0)};DockPanel.SetDock(actions,Dock.Bottom);root.Children.Add(actions);
            var content=new StackPanel();root.Children.Add(new ScrollViewer{Content=content,VerticalScrollBarVisibility=ScrollBarVisibility.Auto});
            Action<string> text=value=>content.Children.Add(new TextBlock{Text=value,TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,10)});
            text("Высота читается из геометрии. Для переменной высоты требуется явное уточнение. Площадь надстройки и площади для доли кровли читаются из расчётных 2D-областей: после расчёта их можно открыть и изменить во вкладке «Результат / Визуальный контроль», затем пересчитать.");
            string h=review.HeightMin.HasValue?review.HeightMin.Value.ToString("0.###")+" - "+review.HeightMax?.ToString("0.###")+" м":"не определена";
            text("Высота: "+h+". Площадь надстройки: "+(review.AnnexArea?.ToString("0.###")??"не определена")+" м².");
            if(row.RoofRatioRequired)text("Числитель - технические помещения: "+(review.TechnicalArea?.ToString("0.###")??"не определён")+" м². Знаменатель - кровля: "+(review.RoofArea?.ToString("0.###")??"не определён")+" м². Доля: "+(review.Ratio?.ToString("0.##%")??"не определена")+".");
            var warnings=new TextBlock{Text=string.Join("\n",review.Errors.Select(p=>(p.Key=="height"?"Высота":p.Key=="area"?"Надстройка":"Кровля")+": "+p.Value)),Foreground=ErrorTextBrush,TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,12)};content.Children.Add(warnings);
            var inputs=new Dictionary<string,Tuple<ComboBox,TextBox>>();
            foreach(var key in new[]{"height","area","ratio"})
            {
                if(!(key=="height"?row.HeightRequired:key=="area"?row.RoofAreaRequired:row.RoofRatioRequired))continue;
                var panel=new WrapPanel{Margin=new Thickness(0,0,0,10)};content.Children.Add(panel);
                panel.Children.Add(new TextBlock{Text=key=="height"?"Высота, м":key=="area"?"Площадь надстройки, м²":"Доля кровли (0-1)",Width=215,VerticalAlignment=VerticalAlignment.Center});
                string mode=key=="height"?row.HeightMode:key=="area"?row.RoofAreaMode:row.RoofRatioMode;
                var choose=new ComboBox{ItemsSource=new[]{"По модели","Вручную"},SelectedIndex=mode=="manual"?1:0,Width=145,Margin=new Thickness(0,0,12,0),ToolTip="По модели - использовать геометрию и расчётные области. Вручную - применить указанное число с предупреждением в отчёте."};panel.Children.Add(choose);
                var input=new TextBox{Width=140,Text=key=="height"?row.Height:key=="area"?row.RoofArea:row.RoofRatio,Padding=new Thickness(5),Visibility=choose.SelectedIndex==1?Visibility.Visible:Visibility.Collapsed,
                    ToolTip="Явное уточнение вместо геометрического результата. Влияет на включение этажа в этажность и количество этажей.\nПример: "+(key=="height"?"2,4":key=="area"?"12,5":"0,15")};panel.Children.Add(input);
                choose.SelectionChanged+=(s,e)=>input.Visibility=choose.SelectedIndex==1?Visibility.Visible:Visibility.Collapsed;inputs[key]=Tuple.Create(choose,input);
            }
            ListBox roofs=null;
            if(row.RoofRatioRequired)
            {
                text("Кровля, к которой относится надстройка. Можно выбрать несколько частей; их пересечения учитываются один раз. Без выбора используется только единственный однозначный вариант.");
                roofs=new ListBox{ItemsSource=review.Roofs,DisplayMemberPath="Display",SelectionMode=SelectionMode.Multiple,Height=150,ToolTip="Выберите связанные части кровли. Не включайте крышу самой надстройки или соседнего здания."};
                foreach(var roof in review.Roofs.Where(r=>(row.RoofSources??new List<string>()).Contains(r.Key)))roofs.SelectedItems.Add(roof);
                content.Children.Add(roofs);
            }
            text("Источники геометрии:\n"+string.Join("\n",review.Details));
            var apply=new Button{Content="Применить",Padding=new Thickness(14,7,14,7),Margin=new Thickness(0,0,10,0),ToolTip="Применить способы расчёта и выбор кровли. Неопределённые данные будут указаны в предварительной проверке."};actions.Children.Add(apply);
            apply.Click+=(s,e)=>
            {
                foreach(var pair in inputs.Where(p=>p.Value.Item1.SelectedIndex==1))
                {double value;if(!TEP.Engine.TryNumber(pair.Value.Item2.Text,out value)||value<0||pair.Key!="ratio"&&value==0||pair.Key=="ratio"&&value>1)
                    {MessageBox.Show(dialog,"Проверьте ручное значение: положительная высота и площадь; доля от 0 до 1.","KPLN | Геометрия этажа");return;}}
                foreach(var pair in inputs)
                {
                    string mode=pair.Value.Item1.SelectedIndex==1?"manual":"auto",value=pair.Value.Item2.Text;
                    if(pair.Key=="height"){row.HeightMode=mode;row.Height=value;}
                    else if(pair.Key=="area"){row.RoofAreaMode=mode;row.RoofArea=value;}
                    else{row.RoofRatioMode=mode;row.RoofRatio=value;}
                }
                if(roofs!=null)row.RoofSources=roofs.SelectedItems.Cast<TEP.LevelRoofOption>().Select(r=>r.Key).ToList();
                dialog.DialogResult=true;
            };
            var cancel=new Button{Content="Закрыть",Padding=new Thickness(14,7,14,7),ToolTip="Закрыть без изменения настроек."};cancel.Click+=(s,e)=>dialog.Close();actions.Children.Add(cancel);
            return dialog;
        }
    }
}
