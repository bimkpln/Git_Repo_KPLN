using System;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    public partial class SS_Sysetm_SelectIncludeElemAlg : Window
    {
        public SS_Sysetm_SelectIncludeElemAlg(string title, string mainTxt)
        {
            InitializeComponent();

            this.Title = $"KPLN: {title}";
            this.MainLable.Text = mainTxt;
        }

        public string ClickedBtnContent { get; private set; }

        public void AddCustomBtn(string buttonText)
        {
            Button btn = new Button
            {
                Content = buttonText,
                Style = this.FindResource("MainButtons") as Style
            };

            btn.Click += Button_Click;

            this.MainStackPanel.Children.Add(btn);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Пример обработки события клика по кнопке
            Button clickedButton = sender as Button;
            ClickedBtnContent = clickedButton.Content.ToString();
            
            DialogResult = true;
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}
