﻿using System.IO;
using System.Reflection;
using static KPLN_ParamSetter.Common.Collections;

namespace KPLN_ParamSetter.Source
{
    public class Source
    {
        public string Value { get; }
        private static string AssemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        public Source(Icon icon)
        {
            switch (icon)
            {
                case Icon.ParamSetter:
                    Value = Path.Combine(AssemblyPath, @"Source\param_setter.png");
                    break;
            }
        }
    }
}
