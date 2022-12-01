using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Classificator.UserInfo
{
    public class StandartUserInfo : UserInfo
    {
        public string getUserName()
        {
            return string.Format(Environment.UserName);
        }
    }
}
