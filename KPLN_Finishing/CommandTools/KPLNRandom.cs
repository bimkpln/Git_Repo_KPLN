using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Finishing.CommandTools
{
    public class KPLNRandom
    {
        private Random random = new Random();
        private List<int> queue = new List<int>();
        private int rvar = 0;
        public KPLNRandom(int c)
        {
            rvar = c;
            random = new Random(rvar);
        }
        public int GetRandom()
        {
            if (queue.Count > 200)
            {
                random = new Random(rvar + (int)DateTime.Now.Ticks);
                int r = random.Next(0, 256);
                queue.Clear();
                queue.Add(r);
                return r;
            }
            else
            {
                int r = random.Next(0, 256);
                while (queue.Contains(r))
                {
                    r = random.Next(0, 256);
                }
                queue.Add(r);
                return r;
            }
        }
    }
}
