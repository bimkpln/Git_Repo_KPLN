using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_User.Common
{
    internal class CheckMEPHeightManager
    {
        public int MainElevation;
        public List<CheckMEPHeightARRoomData> CheckMEPHeightARRoomDataColl;
        public List<CheckMEPHeightARElemData> CheckMEPHeightARElemDataColl;

        public CheckMEPHeightManager(int elevation)
        {
            MainElevation = elevation;
        }

        public void InsertDataByElevation_ARRoomData(int elevation, List<CheckMEPHeightARRoomData> arRoomDataColl)
        {

        }

        public void InsertDataByElevation_ARElemData(int elevation, List<CheckMEPHeightARElemData> arElemDataColl)
        {

        }

        public CheckMEPHeightManager GetARDataByElevation(int elevation)
        {
            return null;
        }
    }
}
