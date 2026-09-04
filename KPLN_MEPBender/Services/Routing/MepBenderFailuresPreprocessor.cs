using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Routing
{
    internal sealed class MepBenderFailuresPreprocessor : IFailuresPreprocessor
    {
        public MepBenderFailuresPreprocessor()
        {
            Failures = new List<MepBenderFailure>();
        }

        public List<MepBenderFailure> Failures { get; }

        public bool HasError { get; private set; }

        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> messages = failuresAccessor.GetFailureMessages();
            if (messages.Count == 0)
                return FailureProcessingResult.Continue;

            bool shouldRollBack = false;
            foreach (FailureMessageAccessor message in messages)
            {
                string description = message.GetDescriptionText();
                Failures.Add(new MepBenderFailure(message.GetFailureDefinitionId(), description));

                if (message.GetSeverity() != FailureSeverity.Warning || IsCriticalForBending(description))
                {
                    shouldRollBack = true;
                    HasError = true;
                    continue;
                }

                TryDeleteWarning(failuresAccessor, message);
            }

            return shouldRollBack
                ? FailureProcessingResult.ProceedWithRollBack
                : FailureProcessingResult.Continue;
        }

        private bool IsCriticalForBending(string description)
        {
            string text = (description ?? string.Empty).ToLowerInvariant();
            return text.Contains("линия слишком коротка")
                   || text.Contains("line is too short")
                   || text.Contains("too short")
                   || text.Contains("недостаточно места")
                   || text.Contains("недостаточно пространства");
        }

        private void TryDeleteWarning(FailuresAccessor failuresAccessor, FailureMessageAccessor message)
        {
            try
            {
                failuresAccessor.DeleteWarning(message);
            }
            catch
            {
                // Same idea as KPLN_Library_OpenDocHandler: failure UI should not block the command.
            }
        }
    }
}