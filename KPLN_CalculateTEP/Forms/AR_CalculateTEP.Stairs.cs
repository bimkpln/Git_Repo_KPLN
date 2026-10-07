using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private void Stairs_Click(object sender,RoutedEventArgs e)
        {
            Try(()=>{var dialog=CreateStairsDialog();dialog.Owner=this;dialog.ShowDialog();});
        }

        private Window CreateStairsDialog()
        {
            var draft=(Config.StairFamilies??new List<TEP.StairFamilySetting>()).Select(s=>new TEP.StairFamilySetting{
                Category=s.Category,Family=s.Family,Type=s.Type,WholeStair=s.WholeStair,WidthParameter=s.WidthParameter,WidthAxis=s.WidthAxis}).ToList();
            var window=new Window{Title="KPLN | Лестницы и Марши",Width=880,Height=560,MinWidth=650,MinHeight=380,WindowStartupLocation=WindowStartupLocation.CenterOwner};
            var root=new DockPanel{Margin=new Thickness(16)};window.Content=root;
            var footer=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right};DockPanel.SetDock(footer,Dock.Bottom);root.Children.Add(footer);
            var apply=new Button{Content="Применить",Padding=new Thickness(18,7,18,7),Margin=new Thickness(6)};footer.Children.Add(apply);
            var cancel=new Button{Content="Отмена",IsCancel=true,Padding=new Thickness(18,7,18,7),Margin=new Thickness(6)};footer.Children.Add(cancel);
            var body=new StackPanel();root.Children.Add(new ScrollViewer{Content=body,VerticalScrollBarVisibility=ScrollBarVisibility.Auto});
            body.Children.Add(CategoryLabel("Стандартные лестницы Revit: ширина каждого марша читается автоматически. Ниже выберите дополнительные типы загружаемых семейств."));
            var tags=new WrapPanel{Margin=new Thickness(0,8,0,12)};body.Children.Add(tags);
            var addRow=new DockPanel();body.Children.Add(addRow);
            var add=new Button{Content="+",Width=36,Margin=new Thickness(8,0,0,0)};DockPanel.SetDock(add,Dock.Right);addRow.Children.Add(add);
            var picker=new ComboBox{ItemsSource=Engine.FamilyChoices,DisplayMemberPath="Display",IsTextSearchEnabled=true};addRow.Children.Add(picker);
            body.Children.Add(CategoryLabel("Выберите тег для настройки. Для отдельного марша используется ширина прямоугольной проекции либо указанный параметр.\nДля целой лестницы - выбранные вложенные марши либо параметр именно ширины марша, не общего габарита."));
            var editor=new StackPanel{Margin=new Thickness(0,8,0,0)};body.Children.Add(editor);
            Action<TEP.StairFamilySetting> edit=setting=>
            {
                editor.Children.Clear();if(setting==null)return;
                editor.Children.Add(CategoryLabel(setting.Caption));
                var mode=new ComboBox{ItemsSource=new[]{"Отдельный марш","Лестница целиком"},SelectedIndex=setting.WholeStair?1:0,Margin=new Thickness(0,4,0,12)};editor.Children.Add(mode);
                mode.SelectionChanged+=(s,args)=>setting.WholeStair=mode.SelectedIndex==1;
                editor.Children.Add(CategoryLabel("Направление ширины отдельного марша в локальных осях семейства"));
                var axis=new ComboBox{ItemsSource=new[]{"Выберите для расчёта по геометрии","X","Y"},SelectedIndex=setting.WidthAxis=="X"?1:setting.WidthAxis=="Y"?2:0,Margin=new Thickness(0,4,0,8)};editor.Children.Add(axis);
                axis.SelectionChanged+=(s,args)=>setting.WidthAxis=axis.SelectedIndex==1?"X":axis.SelectedIndex==2?"Y":"";
                editor.Children.Add(CategoryLabel("Параметр ширины марша (необязательно)"));
                var parameter=new TextBox{Text=setting.WidthParameter??"",Padding=new Thickness(6),ToolTip="Параметр экземпляра или типа: Длина либо текст в метрах. Пусто - геометрия отдельного марша или выбранные вложенные марши."};editor.Children.Add(parameter);
                parameter.TextChanged+=(s,args)=>setting.WidthParameter=parameter.Text.Trim();
            };
            Action refresh=null;
            refresh=()=>
            {
                tags.Children.Clear();
                foreach(var setting in draft)
                {
                    var row=new StackPanel{Orientation=Orientation.Horizontal};
                    var tag=new Border{Background=new SolidColorBrush(Color.FromRgb(231,239,248)),CornerRadius=new CornerRadius(6),Margin=new Thickness(0,0,6,6),Padding=new Thickness(4),Child=row};
                    var select=new Button{Content=setting.Family+" / "+setting.Type,ToolTip=setting.Caption,Padding=new Thickness(6,3,6,3),Background=Brushes.Transparent,BorderThickness=new Thickness(0)};
                    select.Click+=(s,args)=>edit(setting);row.Children.Add(select);
                    var remove=new Button{Content="×",Padding=new Thickness(6,1,6,1),ToolTip="Удалить тип из расчёта ширины марша"};
                    remove.Click+=(s,args)=>{draft.Remove(setting);refresh();edit(draft.FirstOrDefault());};row.Children.Add(remove);tags.Children.Add(tag);
                }
            };
            add.Click+=(s,args)=>
            {
                var selected=picker.SelectedItem as TEP.FamilyChoice;if(selected==null)return;
                var setting=draft.FirstOrDefault(v=>v.Caption==selected.Caption);
                if(setting==null){setting=new TEP.StairFamilySetting{Category=selected.Category,Family=selected.Family,Type=selected.Type};draft.Add(setting);}
                refresh();edit(setting);
            };
            apply.Click+=(s,args)=>{Config.StairFamilies=draft;checkedConfiguration=null;window.DialogResult=true;};
            refresh();edit(draft.FirstOrDefault());return window;
        }
    }
}
