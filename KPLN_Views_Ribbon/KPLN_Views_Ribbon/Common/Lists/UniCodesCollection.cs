using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Views_Ribbon.Common.Lists
{
    internal class UniCodesCollection
    {
        private List<UniEntity> _uniCodes = new List<UniEntity>()
        {
            new UniEntity()
            {
                Name = "RLM",
                Code = "‏"
            },
            new UniEntity()
            { 
                Name = "LRE",
                Code = "‪"
            },
            new UniEntity()
            {
                Name = "LRO",
                Code = "‭"
            },
            new UniEntity()
            {
                Name = "RLO",
                Code = "‮"
            },
            new UniEntity()
            {
                Name = "PDF",
                Code = "‬"
            },
            new UniEntity()
            {
                Name = "RS",
                Code = ""
            },
            new UniEntity()
            {
                Name = "US",
                Code = ""
            }
        };
        
        public List<UniEntity> UniCodes
        {
            get { return _uniCodes; }
        }
    }

    internal sealed class UniEntity
    {
        public string Name { get; set; }
        public string Code { get; set; }
    }
}
