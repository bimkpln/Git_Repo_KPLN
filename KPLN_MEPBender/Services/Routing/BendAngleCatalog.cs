using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Routing
{
    public static class BendAngleCatalog
    {
        private static readonly double[] UserAngles = { 90, 45, 30, 15 };

        public static IReadOnlyCollection<double> GetUserAngles() => UserAngles;
    }
}
