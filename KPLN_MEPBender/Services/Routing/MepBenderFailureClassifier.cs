using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_MEPBender.Services.Routing
{
    internal static class MepBenderFailureClassifier
    {
        public static MepBenderFailureKind Classify(IEnumerable<MepBenderFailure> failures, MepBendResult result, Exception exception)
        {
            string text = BuildSearchText(failures, result, exception);
            if (string.IsNullOrWhiteSpace(text))
                return MepBenderFailureKind.None;

            if (ContainsAny(text,
                "недостаточно места",
                "не достаточно места",
                "недостаточно пространства",
                "линия слишком коротка",
                "направление воздуховода/трубы изменено на противоположное",
                "направление короба изменено на противоположное",
                "направление кабельного лотка изменено на противоположное",
                "изменено на противоположное, что вызвало ошибки соединений",
                "line is too short",
                "not enough space",
                "not enough room",
                "not enough length",
                "too short",
                "insufficient space",
                "insufficient room"))
                return MepBenderFailureKind.InsufficientSpace;

            if (ContainsAny(text,
                "отвод",
                "соединительн",
                "соединить",
                "соединения",
                "фасонн",
                "elbow",
                "fitting",
                "no auto-route solution",
                "cannot find",
                "failed to insert"))
                return MepBenderFailureKind.InvalidFittingFamily;

            return MepBenderFailureKind.Other;
        }

        private static string BuildSearchText(IEnumerable<MepBenderFailure> failures, MepBendResult result, Exception exception)
        {
            List<string> parts = new List<string>();

            if (failures != null)
                parts.AddRange(failures.Select(f => f.Description));

            if (result != null)
                parts.AddRange(result.Issues.Select(i => i.Message));

            if (exception != null)
                parts.Add(exception.ToString());

            return string.Join(" ", parts.Where(p => !string.IsNullOrWhiteSpace(p))).ToLowerInvariant();
        }

        private static bool ContainsAny(string text, params string[] parts)
        {
            return parts.Any(p => text.Contains(p));
        }
    }
}