using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.Core
{
    internal sealed class OpeningConfirmationResult
    {
        internal int Total { get; set; }
        internal int Confirmed { get; set; }
        internal int AlreadyConfirmed { get; set; }
        internal Dictionary<string, string> Errors { get; } = new Dictionary<string, string>();

        internal string Summary => "Все отверстия модели: " + Total
            + ". Подтверждено сейчас: " + Confirmed
            + ". Уже подтверждены: " + AlreadyConfirmed
            + ". Не подтверждено: " + Errors.Count + ".";
    }
}
