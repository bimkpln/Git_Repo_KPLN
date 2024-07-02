using System.Windows;

namespace KPLN_HoleManager.Forms
{
    internal class StringResource : Freezable
    {
        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register("Value", typeof(string), typeof(StringResource));

        public string Value
        {
            get { return (string)GetValue(ValueProperty); }
            set { SetValue(ValueProperty, value); }
        }

        protected override Freezable CreateInstanceCore()
        {
            return new StringResource();
        }
    }
}