using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_User.Common
{
    public static class Collections
    {
        public enum CalculateType { Default, Floor }
        public enum LevelCheckResult { FullyInside, MostlyInside, TheLeastInside, NotInside }
        public enum Status { AllmostOk, LittleWarning, Warning, Error }
        public enum StatusExtended { Critical, Warning, LittleWarning }
        public enum CheckResult { NoSections, Sections, Corpus, Error }
    }
}
