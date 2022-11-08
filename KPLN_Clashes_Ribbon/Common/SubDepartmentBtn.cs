﻿using KPLN_Clashes_Ribbon.Common.Reports;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace KPLN_Clashes_Ribbon.Common
{
    /// <summary>
    /// Коллекция кнопок основных подотделов КПЛН
    /// </summary>
    public sealed class SubDepartmentBtn : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private Brush _delegateBtnBackground;

        public SubDepartmentBtn(int id, string name, string description)
        {
            Id = id;
            Name = name;
            Description = description;
        }

        public int Id { get; private set; }

        public string Name { get; private set; }

        public string Description { get; private set; }
        
        /// <summary>
        /// Привязка к экземпляру отчета
        /// </summary>
        public ReportInstance Parent { get; private set; }

        public Brush DelegateBtnBackground
        {
            get { return _delegateBtnBackground; }
            set 
            { 
                _delegateBtnBackground = value;
                NotifyPropertyChanged();
            }
        }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Установка привязок к инстансам отчетов и установка цвета
        /// </summary>
        public void SetBinding(ReportInstance reportInstance, Brush delegateBtnBackground)
        {
            this.Parent = reportInstance;
            this.DelegateBtnBackground = delegateBtnBackground;
        }
    }
}
