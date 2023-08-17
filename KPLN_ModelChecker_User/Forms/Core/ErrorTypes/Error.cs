namespace KPLN_ModelChecker_User.Forms.Core.ErrorTypes
{
    public sealed class Error : IError
    {
        public static IError Instance => _instance ?? (_instance = new Error());

        public int Id => 1;

        public string Name => "Error";

        public string Description => "Критическая ошибка";

        private static IError _instance;

        public Error()
        {
        }
    }
}
