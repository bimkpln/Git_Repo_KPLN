using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KPLN_RevitWorksets.Common.Collections;

namespace KPLN_RevitWorksets.Source
{
    public class Source
    {
        public string Value { get; }
        private static string AssemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        public Source(Icon icon)
        {
            switch (icon)
            {
                case Icon.Command_large:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\Command_large.png");
                    break;
                default:
                    throw new Exception("Undefined icon!");
            }
        }
    }

}
