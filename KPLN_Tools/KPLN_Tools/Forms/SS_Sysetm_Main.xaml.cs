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
            ICommand executeAddToConsistSystem,
            ICommand executeCreateParallelSystem,
            ICommand executeRefreshSystemCommand)
        {
            CurrentSystemViewEntity = systemViewEntity;
            ExecuteAddParamsCommand = executeAddParamsCommand;
            ExecuteCreateConsistSystemCommand = executeCreateConsistSystem;
            ExecuteAddToConsistSystem = executeAddToConsistSystem;
            ExecuteCreateParallelSystemCommand = executeCreateParallelSystem;
            ExecuteRefreshSystemCommand = executeRefreshSystemCommand;

            InitializeComponent();

            DataContext = CurrentSystemViewEntity;
        }

        public SS_SystemViewEntity CurrentSystemViewEntity { get; set; }

        /// <summary>
        /// Комманда по добвалению параметров в проект
        /// </summary>
        public ICommand ExecuteAddParamsCommand { get; private set; }

        /// <summary>
        /// Комманда по созданию последовательной цепи СС
        /// </summary>
        public ICommand ExecuteCreateConsistSystemCommand { get; private set; }

        /// <summary>
        /// Комманда по ДОБАВЛЕНИЮ элемента к последовательной цепи СС
        /// </summary>
        public ICommand ExecuteAddToConsistSystem { get; private set; }

        /// <summary>
        /// Комманда по созданию параллельной цепи СС
        /// </summary>
        public ICommand ExecuteCreateParallelSystemCommand { get; private set; }

        /// <summary>
        /// Команда по переадресации после внесения измов (без добавления или удаления элементов)
        /// </summary>
        public ICommand ExecuteRefreshSystemCommand { get; private set; }

        private void OnMenuAddParamsBtn_Click(object sender, RoutedEventArgs e)
        {
            Command_SS_Systems.OnIdling_ICommandQueue.Enqueue(ExecuteAddParamsCommand);
        }

        private void OnConsistentlySys_Click(object sender, RoutedEventArgs e)
        {
            Command_SS_Systems.OnIdling_ICommandQueue.Enqueue(ExecuteCreateConsistSystemCommand);
        }

        private void OnConsistentlySysAdd_Click(object sender, RoutedEventArgs e)
        {
            Command_SS_Systems.OnIdling_ICommandQueue.Enqueue(ExecuteAddToConsistSystem);
        }

        private void OnParallelSys_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnParallelSysAdd_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnReculcSys_Click(object sender, RoutedEventArgs e)
        {

        }

        private void RefreshSystem_Click(object sender, RoutedEventArgs e)
        {
            Command_SS_Systems.OnIdling_ICommandQueue.Enqueue(ExecuteRefreshSystemCommand);
        }

        private void RefreshSystemByUser_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnHelp_Click(object sender, RoutedEventArgs e) => System.Diagnostics.Process.Start(@"http://moodle/mod/book/view.php?id=502&chapterid=685");

        private void OnExit_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}
