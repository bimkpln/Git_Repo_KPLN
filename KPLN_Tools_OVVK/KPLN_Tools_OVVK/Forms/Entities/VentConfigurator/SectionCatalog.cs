using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class SectionCatalog : Dictionary<int, SectionDefinition>
    {
        internal string SourceTypeName;
        internal double InstallationWidthMm, InstallationHeightMm, FrameHeightMm;
        internal int IntermediateCount;
        internal bool HasValve;
        internal Dictionary<string, double> SharedLengthsMm = new Dictionary<string, double>();
        internal Dictionary<string, bool?> SharedFlags = new Dictionary<string, bool?>();
        internal FamilyDependencies Dependencies = new FamilyDependencies();

        internal SectionCatalog Copy()
        {
            var result = new SectionCatalog
            {
                SourceTypeName = SourceTypeName,
                InstallationWidthMm = InstallationWidthMm,
                InstallationHeightMm = InstallationHeightMm,
                FrameHeightMm = FrameHeightMm,
                IntermediateCount = IntermediateCount,
                HasValve = HasValve,
                Dependencies = Dependencies,
                SharedLengthsMm = new Dictionary<string, double>(SharedLengthsMm),
                SharedFlags = new Dictionary<string, bool?>(SharedFlags)
            };
            foreach (var pair in this)
                result.Add(pair.Key, new SectionDefinition
                {
                    ParameterName = pair.Value.ParameterName,
                    DisplayName = pair.Value.DisplayName,
                    Choices = pair.Value.Choices,
                    SourceType = pair.Value.SourceType,
                    DefaultLengthMm = pair.Value.DefaultLengthMm,
                    DefaultWidthMm = pair.Value.DefaultWidthMm,
                    DefaultHeightMm = pair.Value.DefaultHeightMm,
                    DefaultOffsetXMm = pair.Value.DefaultOffsetXMm,
                    DefaultOffsetYMm = pair.Value.DefaultOffsetYMm
                });
            return result;
        }

        internal DimensionValue Dimension(string name, double millimeters)
        {
            return new DimensionValue(millimeters, Dependencies.IsCalculated(name), FamilyParameterNames.IsOffset(name));
        }
    }
}
