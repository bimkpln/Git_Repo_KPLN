using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Classificator.UserInfo
{
    public class KplnUserInfo : UserInfo
    {
#if Revit2020 || Revit2018
        public string getUserName()
        {
            return string.Format("{0} {1}", KPLN_Loader.Preferences.User.Family, KPLN_Loader.Preferences.User.Name);
        }
#endif
#if Revit2020Std || Revit2018Std
        public string getUserName()
        {
            throw new NotImplementedException();
        }
#endif
    }
}
