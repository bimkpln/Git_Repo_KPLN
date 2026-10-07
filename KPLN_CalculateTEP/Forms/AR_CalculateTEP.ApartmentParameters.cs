using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private Window BuildApartmentParameterWindow(TEP.UserDictionary categoryDraft=null)
        {
            var window=new Window{Title="KPLN | Запись параметров квартир",Width=1200,Height=700,MinWidth=900,MinHeight=560,WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=Background,FontFamily=FontFamily,FontSize=14};
            window.Resources.MergedDictionaries.Add(Resources);
            var root=new DockPanel{Margin=new Thickness(20)};
            var header=new StackPanel{Margin=new Thickness(0,0,0,14)};DockPanel.SetDock(header,Dock.Top);root.Children.Add(header);
            var data=TEP.Engine.NormalizeRoomCategoryWrites(Engine.RoomCategoryAssignments(departmentGroups??new List<TEP.DepartmentGroup>(),categoryDraft));
            header.Children.Add(new TextBlock{Text="Помещения для записи: "+data.Count,FontSize=18,FontWeight=FontWeights.SemiBold});
            header.Children.Add(new TextBlock{Text="В каждое помещение записывается название его категории из таблицы. Нераспределённые помещения и категории, считаемые по семействам, пропускаются. Запись доступна только в основную модель; помещения связей будут отмечены при проверке.",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,8,0,0)});
            var footer=new StackPanel{Margin=new Thickness(0,14,0,0)};DockPanel.SetDock(footer,Dock.Bottom);root.Children.Add(footer);
            header.Children.Add(CategoryLabel("Параметр для записи"));
            header.Children.Add(new TextBlock{Tag="room-output-parameter",Text=TEP.Engine.DefaultRoomOutputParameter,FontWeight=FontWeights.SemiBold,Margin=new Thickness(0,0,0,8)});
            var buttons=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,14,0,0)};footer.Children.Add(buttons);
            var close=new Button{Content="Закрыть",IsCancel=true};var preview=new Button{Content="Проверить запись",IsEnabled=data.Count>0,ToolTip="Показывает ID помещений, старое и новое значения перед записью в модель."};buttons.Children.Add(close);buttons.Children.Add(preview);
            var grid=new DataGrid{IsReadOnly=true,ItemsSource=data};
            TextColumn(grid,"Источник","Source",180,true);TextColumn(grid,"ID","Element",85,true);TextColumn(grid,"Номер","Number",85,true);
            TextColumn(grid,"Имя","Name",200,true);TextColumn(grid,"Этаж","Level",110,true);TextColumn(grid,"Будет записано: категория","Category",400,true);root.Children.Add(grid);
            buttons.Children.Insert(0,CreateRoomNavigationButton(grid,window));
            preview.Click+=(s,e)=>
            {
                try
                {
                    ShowRoomParameterWrite(window,data);
                    if(Engine.RequestedDetail!=null)CloseNavigationDialog(window);
                }
                catch(Exception ex){MessageBox.Show(window,ex.Message,"KPLN | Запись параметров квартир",MessageBoxButton.OK,MessageBoxImage.Warning);}
            };
            window.Content=root;return window;
        }
    }
}
