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
using Autodesk.Revit.DB;


namespace KPLN_Tools.Forms
{
    public enum Algoritm
    {
        Close,
        SaveSize,
        RevalueSize,
        LoadParams
    }

    /// <summary>
    /// Логика взаимодействия для UserSelectAlgoritm.xaml
    /// </summary>
    public partial class UserSelectAlgoritm : Window
    {
        public Algoritm UserAlgoritm;

        public UserSelectAlgoritm()
        {
            InitializeComponent();
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            UserAlgoritm = Algoritm.Close;
            this.Close();
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            UserAlgoritm = Algoritm.SaveSize;
            this.Close();
        }

        private void OnRevalue(object sender, RoutedEventArgs e)
        {
            UserAlgoritm = Algoritm.RevalueSize;
            this.Close();
        }

        private void OnParams(object sender, RoutedEventArgs e)
        {
            UserAlgoritm = Algoritm.LoadParams;
            this.Close();
        }
    }
}
