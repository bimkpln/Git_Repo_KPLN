using System;
using System.Windows.Input;

namespace KPLN_Tools.Common.SS_System
{
    public class SS_SystemCommand : ICommand
    {
        private Action _execute;
        public event EventHandler CanExecuteChanged;

        public SS_SystemCommand(Action execute)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        }

        public bool CanExecute(object parameter)
        {
            // Можете добавить логику для определения, может ли команда быть выполнена в данный момент
            return true;
        }

        public void Execute(object parameter)
        {
            _execute();
        }
    }
}
