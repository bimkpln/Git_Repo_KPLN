using KPLN_ModelChecker_User.Forms;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_User.WPFItems
{
    /// <summary>
    /// Спец. класс-обертка, для подготовки данных для передачи в отчет
    /// </summary>
    public sealed class WPFReportCreator
    {
        public WPFReportCreator(IEnumerable<WPFEntity> wpfEntityColl, IEnumerable<WPFEntity> wpfFiltration, string checkName, string logLastRun, string logMarker)
        {
            WPFEntityCollection = wpfEntityColl.OrderByDescending(w => (int)w.CurrentStatus).ToList();
            int counter = 0;
            foreach (WPFEntity w in WPFEntityCollection)
                w.Header = $"#{++counter} {w.Header}";

            if (wpfFiltration != null)
                WPFFiltration.AddRange(wpfFiltration.OrderBy(w => w.Name).ToList());
            
            CheckName = checkName;
            LogLastRun = logLastRun;
            LogMarker = logMarker;
        }

        public List<WPFEntity> WPFEntityCollection { get; private set; }

        public List<WPFEntity> WPFFiltration { get; private set; } = new List<WPFEntity>() { new WPFEntity("<Все элементы>") };

        public string CheckName { get; }

        public string LogLastRun { get; }

        public string LogMarker { get; }
    }
}
