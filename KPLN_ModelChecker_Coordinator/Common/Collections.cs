namespace KPLN_ModelChecker_Coordinator.Common
{
    public static class Collections
    {
        public enum CalculateType { Default, Floor }
        public enum WarningStatus { Opened, ReOpened, Closed, Ignore, RequestToIgnore, RejectedToIgnore, Removed }
        public enum Icon { OpenManager, Statistics, Preferences }
        public enum LevelCheckResult { FullyInside, MostlyInside, TheLeastInside, NotInside }
        public enum CheckResult { NoSections, Sections, Corpus, Error }
        public enum ElementStatus { NewElement, Approved, Rejected, Removed, Changed }
    }
}
