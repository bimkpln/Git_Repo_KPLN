namespace KPLN_OpeningHoleManager.Core
{
    /// <summary>
    /// Копия настроек отступов из KPLN_Clashes_Ribbon.Core.ZoomSettings для чтения того же локального конфига.
    /// </summary>
    public sealed class OpeningZoomSettings
    {
        // Сохраняем масштаб исходного ZoomSettings, чтобы зум совпадал с Clashes.
        internal const double FeetToMillimeters = 300;
        internal const double DefaultOffsetXYMillimeters = 1500;
        internal const double DefaultOffsetZMinMillimeters = 600;
        internal const double DefaultOffsetZMaxMillimeters = 300;

        public double OffsetXMillimeters { get; set; } = DefaultOffsetXYMillimeters;
        public double OffsetYMillimeters { get; set; } = DefaultOffsetXYMillimeters;
        public double OffsetZMinMillimeters { get; set; } = DefaultOffsetZMinMillimeters;
        public double OffsetZMaxMillimeters { get; set; } = DefaultOffsetZMaxMillimeters;
    }
}
