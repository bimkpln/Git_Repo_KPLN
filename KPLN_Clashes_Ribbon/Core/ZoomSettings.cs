using KPLN_Library_ConfigWorker.Core;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Clashes_Ribbon.Core
{
    public sealed class ZoomSettings : INotifyPropertyChanged, IJsonSerializable
    {
        public const double FeetToMillimeters = 300;
        public const double DefaultOffsetXMillimeters = 1500;
        public const double DefaultOffsetYMillimeters = 1500;
        public const double DefaultOffsetZMinMillimeters = 600;
        public const double DefaultOffsetZMaxMillimeters = 300;
        public const bool DefaultFitCollision = false;

        private double _offsetXMillimeters = DefaultOffsetXMillimeters;
        private double _offsetYMillimeters = DefaultOffsetYMillimeters;
        private double _offsetZMinMillimeters = DefaultOffsetZMinMillimeters;
        private double _offsetZMaxMillimeters = DefaultOffsetZMaxMillimeters;
        private bool _fitCollision = DefaultFitCollision;

        public event PropertyChangedEventHandler PropertyChanged;

        public double OffsetXMillimeters { get => _offsetXMillimeters; set => SetValue(ref _offsetXMillimeters, value); }
        public double OffsetYMillimeters { get => _offsetYMillimeters; set => SetValue(ref _offsetYMillimeters, value); }
        public double OffsetZMinMillimeters { get => _offsetZMinMillimeters; set => SetValue(ref _offsetZMinMillimeters, value); }
        public double OffsetZMaxMillimeters { get => _offsetZMaxMillimeters; set => SetValue(ref _offsetZMaxMillimeters, value); }

        public bool FitCollision
        {
            get => _fitCollision;
            set
            {
                if (_fitCollision == value)
                    return;

                _fitCollision = value;
                NotifyPropertyChanged();
            }
        }

        public double OffsetXFeet => OffsetXMillimeters / FeetToMillimeters;
        public double OffsetYFeet => OffsetYMillimeters / FeetToMillimeters;
        public double OffsetZMinFeet => OffsetZMinMillimeters / FeetToMillimeters;
        public double OffsetZMaxFeet => OffsetZMaxMillimeters / FeetToMillimeters;

        public object ToJson()
        {
            return new
            {
                OffsetXMillimeters,
                OffsetYMillimeters,
                OffsetZMinMillimeters,
                OffsetZMaxMillimeters,
                FitCollision,
            };
        }

        public void SetDefaults()
        {
            OffsetXMillimeters = DefaultOffsetXMillimeters;
            OffsetYMillimeters = DefaultOffsetYMillimeters;
            OffsetZMinMillimeters = DefaultOffsetZMinMillimeters;
            OffsetZMaxMillimeters = DefaultOffsetZMaxMillimeters;
            FitCollision = DefaultFitCollision;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ZoomSettings other))
                return false;

            return Math.Abs(OffsetXMillimeters - other.OffsetXMillimeters) < 0.0001
                && Math.Abs(OffsetYMillimeters - other.OffsetYMillimeters) < 0.0001
                && Math.Abs(OffsetZMinMillimeters - other.OffsetZMinMillimeters) < 0.0001
                && Math.Abs(OffsetZMaxMillimeters - other.OffsetZMaxMillimeters) < 0.0001
                && FitCollision == other.FitCollision;
        }

        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 23 + OffsetXMillimeters.GetHashCode();
            hash = hash * 23 + OffsetYMillimeters.GetHashCode();
            hash = hash * 23 + OffsetZMinMillimeters.GetHashCode();
            hash = hash * 23 + OffsetZMaxMillimeters.GetHashCode();
            hash = hash * 23 + FitCollision.GetHashCode();
            return hash;
        }

        private void SetValue(ref double field, double value, [CallerMemberName] string propertyName = "")
        {
            if (Math.Abs(field - value) < 0.0001)
                return;

            field = value;
            NotifyPropertyChanged(propertyName);
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

