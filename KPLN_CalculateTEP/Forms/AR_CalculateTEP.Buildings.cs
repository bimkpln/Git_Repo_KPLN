using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using TEP=KPLN_CalculateTEP.Common.TepCalculation;

namespace KPLN_CalculateTEP.Forms
{
    public partial class AR_CalculateTEP
    {
        private void Buildings_Click(object sender,RoutedEventArgs args)
        {BuildBuildingsWindow().ShowDialog();}
        private Window BuildBuildingsWindow()
        {
            var dialog=new Window{Title="KPLN | Соответствия корпусов",Owner=IsLoaded?this:null,Width=1150,Height=580,MinWidth=900,MinHeight=420,WindowStartupLocation=WindowStartupLocation.CenterOwner,Resources=Resources,Background=Background};
            var root=new DockPanel{Margin=new Thickness(16)};dialog.Content=root;
            var hint=new TextBlock{Text="Назначайте любые названия и любое количество корпусов. Правила используются только при отсутствии корпуса у объекта. Можно сопоставить имя уровня, источника, группы шахт или значение выбранного параметра. По умолчанию сравнение точное, без учёта регистра; «Содержит» включает поиск части текста. Правила разных корпусов не должны пересекаться. Параметры модели не изменяются; соответствия сохраняются вместе с настройками ТЭП в RVT.",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,12)};DockPanel.SetDock(hint,Dock.Top);root.Children.Add(hint);
            var rows=new ObservableCollection<TEP.BuildingAssignment>(TEP.Engine.Deserialize<System.Collections.Generic.List<TEP.BuildingAssignment>>(TEP.Engine.Serialize(Config.BuildingAssignments??new System.Collections.Generic.List<TEP.BuildingAssignment>())));
            var table=new DataGrid{AutoGenerateColumns=false,CanUserAddRows=true,CanUserDeleteRows=true,ItemsSource=rows};
            var choices=new[]{new TEP.Choice("level","Имя уровня",""),new TEP.Choice("group","Имя группы шахт",""),new TEP.Choice("source","Имя источника",""),new TEP.Choice("parameter","Значение параметра"," ")};
            ChoiceColumn(table,"Сопоставлять по","Scope",choices,190);TextColumn(table,"Имя параметра (для значения)","Parameter",230);TextColumn(table,"Искомое значение / имя","Match",250);CheckColumn(table,"Содержит","Contains",85);TextColumn(table,"Назначить корпус","Building",210);
            var buttons=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,12,0,0)};DockPanel.SetDock(buttons,Dock.Bottom);root.Children.Add(buttons);
            var remove=new Button{Content="Удалить строку"};remove.Click+=(s,e)=>{foreach(var row in table.SelectedItems.OfType<TEP.BuildingAssignment>().ToList())rows.Remove(row);};buttons.Children.Add(remove);
            var cancel=new Button{Content="Отмена",IsCancel=true};buttons.Children.Add(cancel);
            var apply=new Button{Content="Применить"};buttons.Children.Add(apply);
            apply.Click+=(s,e)=>
            {
                table.CommitEdit(DataGridEditingUnit.Cell,true);table.CommitEdit(DataGridEditingUnit.Row,true);
                var meaningful=rows.Where(r=>!string.IsNullOrWhiteSpace(r.Match)||!string.IsNullOrWhiteSpace(r.Building)||!string.IsNullOrWhiteSpace(r.Parameter)).ToList();
                if(meaningful.Any(r=>string.IsNullOrWhiteSpace(r.Match)||string.IsNullOrWhiteSpace(r.Building)||r.Scope=="parameter"&&string.IsNullOrWhiteSpace(r.Parameter)))
                {MessageBox.Show(dialog,"Заполните искомое значение и корпус. Для сопоставления по параметру также укажите его имя.",dialog.Title);return;}
                Config.BuildingAssignments=meaningful;checkedConfiguration=null;dialog.DialogResult=true;
            };
            root.Children.Add(table);return dialog;
        }
    }
}
