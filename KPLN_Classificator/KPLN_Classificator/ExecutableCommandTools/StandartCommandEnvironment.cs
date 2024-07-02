using KPLN_Loader.Common;
using System.Collections.Generic;

namespace KPLN_Classificator.ExecutableCommand
{
    public class StandartCommandEnvironment : CommandEnvironment
    {
        private readonly Queue<IExecutableCommand> commandQueue = new Queue<IExecutableCommand>();

        public Queue<IExecutableCommand> getQueue()
        {
            return commandQueue;
        }

        public void toEnqueue(object obj)
        {
            commandQueue.Enqueue(obj as IExecutableCommand);
        }
    }
}
