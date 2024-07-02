using KPLN_Tools.Common.SS_System;
using KPLN_Tools.ExternalCommands;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{

    public partial class SS_Sysetm_Main : Window
    {
        public SS_Sysetm_Main(
            SS_SystemViewEntity systemViewEntity,
            ICommand executeAddParamsCommand,
            ICommand executeCreateConsistSystem,
            ICommand executeCreateParallelSystem)
        {
            CurrentSS_SystemViewEntity = systemViewEntity;
            ExecuteAddParamsCommand = executeAddParamsCommand;
            ExecuteCreateConsistSystem = executeCreateConsistSystem;
            ExecuteCreateParallelSystem = executeCreateParallelSystem;

            InitializeComponent();
            DataContext = CurrentSS_SystemViewEntity;
        }

        public SS_SystemViewEntity CurrentSS_SystemViewEntity { get; set; }

        /// <summary>
        /// Комманда по добвалению параметров в проект
        /// </summary>
        public ICommand ExecuteAddParamsCommand { get; private set; }

        /// <summary>
        /// Комманда по созданию последовательной цепи СС
        /// </summary>
        public ICommand ExecuteCreateConsistSystem { get; private set; }

        /// <summary>
        /// Комманда по созданию параллельной цепи СС
        /// </summary>
        public ICommand ExecuteCreateParallelSystem { get; private set; }

        private void OnMenuAddParamsBtn_Click(object sender, RoutedEventArgs e)
        {
            Command_SS_Systems.OnIdling_ICommandQueue.Enqueue(ExecuteAddParamsCommand);
        }

        private void OnConsistentlySys_Click(object sender, RoutedEventArgs e)
        {
            Command_SS_Systems.OnIdling_ICommandQueue.Enqueue(ExecuteCreateConsistSystem);
        }

        private void OnParallelSys_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnReculcSys_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnReadress_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnIncrementAdress_Clicked(object sender, RoutedEventArgs e)
        {

        }

        private void OnDecrementAdress_Clicked(object sender, RoutedEventArgs e)
        {

        }

        private void OnHelp_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OnNew_Click(object sender, RoutedEventArgs e)
        {

        }


    }
}
