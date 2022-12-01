using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Classificator.ExecutableCommand
{
    public class StandartCommandEnvironment : CommandEnvironment
    {
        private readonly Queue<MyExecutableCommand> commandQueue = new Queue<MyExecutableCommand>();

        public Queue<MyExecutableCommand> getQueue()
        {
            return commandQueue;
        }

        public void toEnqueue(object obj)
        {
            commandQueue.Enqueue(obj as MyExecutableCommand);
        }
    }
}
