using System;
using System.Globalization;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class DimensionValue : EditableValue
    {
        private string _text;
        private string _sourceText;
        private double _sourceMillimeters;

        internal bool AllowsSigned
        { get; private set; }

        public bool IsCalculated
        { get; private set; }

        public bool IsReadOnly
        {
            get
            {
                return IsCalculated;
            }
        }

        public bool IsCurrent { get; private set; } = true;

        public string Text
        {
            get
            {
                return IsCalculated && !IsCurrent ? "—" : _text;
            }
            set
            {
                if (IsReadOnly || _text == value)
                    return;
                _text = value;
                NotifyValue();
            }
        }

        private void NotifyValue()
        {
            Changed(nameof(Text));
            Changed(nameof(IsValid));
            Changed(nameof(Display));
        }

        internal void SetCalculatedMode(bool calculated)
        {
            if (IsCalculated == calculated)
                return;
            IsCalculated = calculated;
            IsCurrent = !calculated;
            Changed(nameof(IsCalculated));
            Changed(nameof(IsReadOnly));
            Changed(nameof(IsCurrent));
            NotifyValue();
        }

        internal void Invalidate()
        {
            if (IsCalculated)
            {
                IsCurrent = false;
                Changed(nameof(IsCurrent));
                NotifyValue();
            }
        }

        internal void ReadResult(double millimeters)
        {
            if (!IsCalculated)
                return;
            ReadSource(millimeters);
            Changed(nameof(IsCurrent));
            NotifyValue();
        }

        private void ReadSource(double millimeters)
        {
            _sourceMillimeters = millimeters;
            _text = _sourceText = IsFinite(millimeters) ? millimeters.ToString("0.######", CultureInfo.CurrentCulture) : string.Empty;
            IsCurrent = IsFinite(millimeters);
        }

        internal void ReadInput(double millimeters)
        {
            if (IsCalculated)
                return;
            ReadSource(millimeters);
            NotifyValue();
        }

        internal static bool IsPositive(double value)
        {
            return IsFinite(value) && value > 0;
        }

        internal static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        internal bool TryValue(out double value)
        {
            return TryNumber(out value) && (AllowsSigned || value > 0);
        }

        internal bool TryNumber(out double value)
        {
            value = 0;
            if (IsCalculated && !IsCurrent)
                return false;
            // Округление в поле ввода не должно превращаться в изменение параметра Revit.
            if (_text == _sourceText)
            {
                value = _sourceMillimeters;
                return IsFinite(value);
            }
            return double.TryParse((_text ?? "").Trim().Replace(',', '.'),
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value)
                && IsFinite(value);
        }

        public bool IsValid
        {
            get
            {
                double value;
                return IsCalculated || TryValue(out value);
            }
        }

        public string Display
        {
            get
            {
                double value;
                return TryNumber(out value) ? value.ToString("0.######", CultureInfo.CurrentCulture) : "—";
            }
        }

        internal DimensionValue(double millimeters, bool calculated = false, bool allowsSigned = false)
        {
            ReadSource(millimeters);
            IsCalculated = calculated;
            AllowsSigned = allowsSigned;
        }

        internal double Millimeters(string name)
        {
            double value;
            if (!TryValue(out value))
                throw new ArgumentException("«" + name + "»: введите число" + (AllowsSigned ? "" : " больше нуля") + " в миллиметрах.");
            return value;
        }

        internal DimensionValue Copy()
        {
            return new DimensionValue(_sourceMillimeters, IsCalculated, AllowsSigned)
            {
                _text = _text,
                _sourceText = _sourceText,
                IsCurrent = IsCurrent
            };
        }
    }
}
