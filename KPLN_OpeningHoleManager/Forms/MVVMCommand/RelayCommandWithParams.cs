using System;
using System.Windows.Input;

namespace KPLN_OpeningHoleManager.Forms.MVVMCommand
{
    public sealed class RelayCommandWithParams : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommandWithParams(Action<object> execute, Func<bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute == null || _canExecute();

        public void Execute(object parameter) => _execute(parameter);

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
