namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class BooleanValue : EditableValue
    {
        private bool? _value;

        public bool IsCalculated
        { get; private set; }

        public bool IsEnabled
        {
            get
            {
                return !IsCalculated;
            }
        }

        public bool IsValid
        {
            get
            {
                return IsCalculated || _value.HasValue;
            }
        }

        public bool? Value
        {
            get
            {
                return _value;
            }
            set
            {
                if (IsCalculated || _value == value)
                    return;
                _value = value;
                NotifyValue();
            }
        }

        public int SideIndex
        {
            get
            {
                return !_value.HasValue ? -1 : _value.Value ? 0 : 1;
            }
            set
            {
                if (value == 0 || value == 1)
                    Value = value == 0;
            }
        }

        private void NotifyValue()
        {
            Changed(nameof(Value));
            Changed(nameof(SideIndex));
            Changed(nameof(IsValid));
        }

        internal BooleanValue(bool? value, bool calculated = false)
        {
            _value = value;
            IsCalculated = calculated;
        }

        internal void SetCalculatedMode(bool calculated)
        {
            if (IsCalculated == calculated)
                return;
            IsCalculated = calculated;
            Changed(nameof(IsCalculated));
            Changed(nameof(IsEnabled));
            Changed(nameof(IsValid));
        }

        internal void ReadResult(bool? value)
        {
            if (_value == value)
                return;
            _value = value;
            NotifyValue();
        }

        internal BooleanValue Copy()
        {
            return new BooleanValue(_value, IsCalculated);
        }
    }
}
