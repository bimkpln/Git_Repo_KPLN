using Autodesk.Revit.DB;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_ExtraFilter.Forms.Entities.SetParamsByFrame;
using KPLN_Library_Forms.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SetParamsByFrameForm : Window
    {
        public SetParamsByFrameForm(IEnumerable<Element> elemsToSet, IEnumerable<ParamEntity> paramsEntities, object lastRunConfigObj)
        {
            InitializeComponent();

            if (lastRunConfigObj != null && lastRunConfigObj is IEnumerable<MainItem> mainEntites)
                CurrentSetParamsByFrameEntity = new SetParamsByFrameEntity(elemsToSet, paramsEntities.OrderBy(ent => ent.RevitParamName), mainEntites);
            else
                CurrentSetParamsByFrameEntity = new SetParamsByFrameEntity(elemsToSet, paramsEntities.OrderBy(ent => ent.RevitParamName));

            DataContext = CurrentSetParamsByFrameEntity;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public SetParamsByFrameEntity CurrentSetParamsByFrameEntity { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            else if (e.Key == Key.Enter)
                RunBtn_Click(sender, e);
        }

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            KPLN_Loader.Application.OnIdling_CommandQueue
                .Enqueue(new SetParamsByFrameExcCmd(CurrentSetParamsByFrameEntity));

            Close();
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            MainItem defaultMI = new MainItem(CurrentSetParamsByFrameEntity.AllParamEntities.FirstOrDefault());
            CurrentSetParamsByFrameEntity.MainItems.Add(defaultMI);
            CurrentSetParamsByFrameEntity.RunButtonContext();

            DataContext = CurrentSetParamsByFrameEntity;
        }

        private void ClearBtn_Click(object sender, RoutedEventArgs e)
        {
            CurrentSetParamsByFrameEntity.MainItems.Clear();
            CurrentSetParamsByFrameEntity.RunButtonContext();

            DataContext = CurrentSetParamsByFrameEntity;
        }

        private void MenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            if (!((MenuItem)e.Source is MenuItem menuItem))
                return;

            if (!(menuItem.DataContext is MainItem entity))
                return;

            UserDialog ud = new UserDialog("ВНИМАНИЕ",
                $"Сейчас будут удален параметр \"{entity.UserSelectedParamEntity.RevitParamName}\". Продолжить?");
            
            
            if((bool)ud.ShowDialog())
            {
                CurrentSetParamsByFrameEntity.MainItems.Remove(entity);
                CurrentSetParamsByFrameEntity.RunButtonContext();
            }

            DataContext = CurrentSetParamsByFrameEntity;
        }
    }
}
