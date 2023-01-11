using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KPLN_Quantificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для OutputForm.xaml
    /// </summary>
    public partial class OutputForm : Window
    {
        private static OutputForm _instance;

        public static OutputForm GetInstance()
        {
            if (_instance == null)
                _instance = new OutputForm();
            return _instance;
        }

        private OutputForm()
        {
            InitializeComponent();
        }
        
        public void AddHeaderTextBlock(string message)
        {
            TextBlock text_block = new TextBlock();
            text_block.Text = message;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            margin.Top = 10;
            margin.Bottom = 5;
            text_block.Margin = margin;
            text_block.FontSize = 15;
            text_block.FontWeight = FontWeights.Black;
            text_block.TextWrapping = TextWrapping.Wrap;
            text_block.BaselineOffset = 2;
            this.stack_panel.Children.Add(text_block);
        }
        
        public void AddTextBlock(string message)
        {
            TextBlock text_block = new TextBlock();
            text_block.Text = message;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            margin.Top = 2;
            margin.Bottom = 2;
            text_block.Margin = margin;
            text_block.TextWrapping = TextWrapping.Wrap;
            text_block.BaselineOffset = 2;
            this.stack_panel.Children.Add(text_block);
        }
        
        public void AddErrorTextBlock(string message)
        {
            TextBlock text_block = new TextBlock();
            text_block.Text = message;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            margin.Top = 2;
            margin.Bottom = 2;
            text_block.Margin = margin;
            text_block.Foreground = new SolidColorBrush(Color.FromArgb(255, 190, 0, 0));
            text_block.Background = new SolidColorBrush(Color.FromArgb(255, 255, 196, 193));
            text_block.TextWrapping = TextWrapping.Wrap;
            text_block.BaselineOffset = 2;
            this.stack_panel.Children.Add(text_block);
        }
        public void AddSuccessTextBlock(string message)
        {
            TextBlock text_block = new TextBlock();
            text_block.Text = message;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            margin.Top = 2;
            margin.Bottom = 2;
            text_block.Margin = margin;
            text_block.Foreground = new SolidColorBrush(Color.FromArgb(255, 60, 140, 0));
            text_block.Background = new SolidColorBrush(Color.FromArgb(255, 203, 255, 174));
            text_block.TextWrapping = TextWrapping.Wrap;
            text_block.BaselineOffset = 2;
            this.stack_panel.Children.Add(text_block);
        }

        private void On_Closed(object sender, System.ComponentModel.CancelEventArgs e) => _instance = null;
    }
}
