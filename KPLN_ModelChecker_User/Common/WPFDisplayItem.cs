﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Common
{
    public class WPFDisplayItem : INotifyPropertyChanged
    {
        public BoundingBoxXYZ Box { get; private set; }
        public XYZ Centroid { get; private set; }
        public int CategoryId { get; }
        public string Icon { get; set; }
        public Element Element { get; private set; }
        public void SetZoomParams(Element element, BoundingBoxXYZ box)
        {
            Element = element;
            if (box != null)
            {
                Box = box;
                Centroid = new XYZ((box.Max.X + box.Min.X) / 2, (box.Max.Y + box.Min.Y) / 2, (box.Max.Z + box.Min.Z) / 2);
            }
        }
        private ObservableCollection<WPFDisplayItem> _collection { get; set; }
        public ObservableCollection<WPFDisplayItem> Collection
        {
            get { return _collection; }
            set
            {
                _collection = value;
                NotifyPropertyChanged();
            }
        }
        private int _elementId { get; set; }
        public int ElementId
        {
            get { return _elementId; }
            set
            {
                _elementId = value;
                NotifyPropertyChanged();
            }
        }
        private string _toolTip { get; set; }
        public string ToolTip
        {
            get { return _toolTip; }
            set
            {
                _toolTip = value;
                NotifyPropertyChanged();
            }
        }
        private string _header { get; set; }
        public string Header
        {
            get { return _header; }
            set
            {
                _header = value;
                NotifyPropertyChanged();
            }
        }
        private string _description { get; set; }
        public string Description
        {
            get { return _description; }
            set
            {
                _description = value;
                NotifyPropertyChanged();
            }
        }
        private string _category { get; set; }
        public string Category
        {
            get { return _category; }
            set
            {
                _category = value;
                NotifyPropertyChanged();
            }
        }
        private string _name { get; set; }
        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }
        private SolidColorBrush _background { get; set; }
        public SolidColorBrush Background
        {
            get { return _background; }
            set
            {
                _background = value;
                NotifyPropertyChanged();
            }
        }
        private bool _isEnabled { get; set; }
        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                _isEnabled = value;
                NotifyPropertyChanged();
            }
        }
        private System.Windows.Visibility _visibility { get; set; }
        public System.Windows.Visibility Visibility 
        {
            get { return _visibility; }
            set 
            {
                _visibility = value;
                NotifyPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public WPFDisplayItem(int categoryId, StatusExtended status, string icon= "🔍")
        {
            Icon = icon;
            CategoryId = categoryId;
            switch (status)
            {
                case StatusExtended.Critical:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 195, 0));
                    break;
                case StatusExtended.Warning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 218, 247, 166));
                    break;
            }
        }

        public WPFDisplayItem(int categoryId, StatusExtended status, int elementId) : this (categoryId, status)
        {
            ElementId = elementId;
        }
    }
}
