﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_User.ExternalCommands;
using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ModelChecker_Batch.Forms.Entities
{
    public sealed class CheckEntity: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isChecked = false;

        public CheckEntity(AbstrCheck abstrCheck)
        {
            CurrentAbstrCheck = abstrCheck;
            Name = abstrCheck.PluginName;
        }

        public string Name { get; }

        public bool IsChecked 
        { 
            get => _isChecked;
            set
            {
                _isChecked = value;
                OnPropertyChanged();
            }
        }

        public AbstrCheck CurrentAbstrCheck { get; }

        public CheckerEntity[] RunCommand(Document doc)
        {
            Element[] elemsToCheck = CurrentAbstrCheck.GetElemsToCheck(doc);
            return CurrentAbstrCheck.ExecuteCheck(doc, elemsToCheck, false);
        }

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
