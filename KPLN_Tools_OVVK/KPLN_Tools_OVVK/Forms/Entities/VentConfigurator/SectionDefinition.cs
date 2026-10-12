using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class SectionDefinition
    {
        internal string ParameterName, DisplayName;
        internal List<SectionTypeChoice> Choices;
        internal SectionTypeChoice SourceType;
        internal double DefaultLengthMm, DefaultWidthMm, DefaultHeightMm, DefaultOffsetXMm, DefaultOffsetYMm;

        internal static List<SectionTypeChoice> OrderChoices(IEnumerable<SectionTypeChoice> choices)
        {
            return choices.GroupBy(c => c.Key).Select(g => g.First()).OrderBy(c => c.IsEmpty ? 0 : 1)
                .ThenBy(c => c.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
        }
    }
}
