namespace KPLN_ModelChecker_Lib.Common.ErrorTypes
{
    public sealed class LittleWarning : IError
    {
        public static IError Instance => _instance ?? (_instance = new LittleWarning());

        public int Id => 3;

        public string Name => "Warning";

        public string Description => "Обрати внимание";

        private static IError _instance;

        public LittleWarning()
        {
        }
    }
}
