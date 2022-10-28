namespace KPLN_ModelChecker_Lib.Common.ErrorTypes
{
    public sealed  class Warning : IError
    {
        public static IError Instance => _instance ?? (_instance = new Warning());

        public int Id => 2;

        public string Name => "Warning";

        public string Description => "Предупреждение";

        private static IError _instance;

        public Warning()
        {
        }
    }
}
