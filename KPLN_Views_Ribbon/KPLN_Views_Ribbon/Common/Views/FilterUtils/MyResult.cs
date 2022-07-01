﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Views_Ribbon.Views.FilterUtils
{
    public enum ResultType { ok, warning, cancel, error}
    public class MyDialogResult
    {
        public ResultType ResultType;
        public string Message;

        public MyDialogResult(ResultType Result, string Message)
        {
            ResultType = Result;
            this.Message = Message;
        }

    }
}
