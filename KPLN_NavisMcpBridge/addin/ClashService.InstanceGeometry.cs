using System.Collections.Generic;
using System.Linq;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace KPLN_NavisMcpBridge
{
    internal static partial class ClashService
    {
        private static List<ComApi.InwOaFragment3> GetInstanceFragments(ComApi.InwOaPath path, out int skipped)
        {
            return GeometryInstancePath.Select(
                Enumerate(path.Fragments()).Cast<ComApi.InwOaFragment3>(),
                GeometryInstancePath.Read(path.ArrayData), fragment => fragment.path.ArrayData, out skipped);
        }
    }
}
