using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Tools.Common.Collections;

namespace KPLN_Tools.Source
{
    public class Source
    {
        public string Value { get; }
        private static string AssemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        public Source(Icon icon)
        {
            switch (icon)
            {
                case Icon.toolBox:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\toolBox.png");
                    break;
                case Icon.pushPin:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\pushPin.png");
                    break;
                case Icon.renamerFunc:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\renamerFunc.png");
                    break;
                case Icon.autonumber:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\autonumber.png");
                    break;
                default:
                    throw new Exception("Undefined icon!");
            }
        }
    }

}
