using KPLN_Loader.Common;

namespace KPLN_Classificator.ExecutableCommand
{
    public class KplnCommandEnvironment : CommandEnvironment
    {
        public void toEnqueue(object obj)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(obj as IExecutableCommand);
        }
    }
}
