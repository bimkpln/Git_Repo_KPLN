namespace KPLN_ModelChecker_User.Common
{
    public class RoomParamData
    {
        /// <summary>
        /// Допуск в 1 м² (double - в футах!) в перделах 1 квартиры
        /// </summary>
        public readonly double SumAreaTolerance = 10.764;

        /// <summary>
        /// Допуск в N символов для разницы в именовании и нумерации
        /// </summary>
        public readonly int TextTolerance = 1;

        public RoomParamData(string firstParamName, string secondParamName)
        {
            FirstParam = firstParamName;
            SecondParam = secondParamName;
        }

        public string FirstParam { get; private set; }

        public string SecondParam { get; private set; }
    }
}
