#if Revit2020 || Revit2018
using KPLN_Loader.Common;
#endif

namespace KPLN_Classificator.ExecutableCommand
{
    public class KplnCommandEnvironment : CommandEnvironment
    {
#if Revit2020 || Revit2018
        public void toEnqueue(object obj)
        {
            KPLN_Loader.Preferences.CommandQueue.Enqueue(obj as IExecutableCommand);
        }
#endif
#if Revit2020Std || Revit2018Std
        public void toEnqueue(object obj)
        {
            throw new System.NotImplementedException();
        }
#endif
    }
}
