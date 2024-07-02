using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;


namespace KPLN_Quantificator.Forms
{
    public partial class RenameViewForm : Window
    {
        Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
        static string viewpointTextFloor, viewpointTextNumber;

        public RenameViewForm()
        {
            InitializeComponent();
            RecordingViewpointName();
        }

        public void RenameViewPointObj()
        {

            var parent = doc.SavedViewpoints.CurrentSavedViewpoint.Parent;
            int index = parent.Children.IndexOf(doc.SavedViewpoints.CurrentSavedViewpoint);

            SavedItem newViewpoint = doc.SavedViewpoints.CurrentSavedViewpoint.CreateUniqueCopy();

            newViewpoint.DisplayName = viewpointTextFloor + "." + viewpointTextNumber + "." + ViewPointTextName.Text;

            if (int.TryParse(viewpointTextNumber, out int number))
            {
                viewpointTextNumber = (number + 1).ToString("000");
            }

            doc.SavedViewpoints.ReplaceWithCopy(parent, index, newViewpoint);
        }


        private void RecordingViewpointName()
        {
            ViewPointTextName.Text = doc.SavedViewpoints.CurrentSavedViewpoint.DisplayName;
            ViewPointTextList.Text = doc.SavedViewpoints.CurrentSavedViewpoint.DisplayName;
        }

        private void Button_CloseClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_SaveClick(object sender, RoutedEventArgs e)
        {
            viewpointTextFloor = ViewPointTextFloor.Text;
            viewpointTextNumber = ViewPointTextNumber.Text;
            this.Close();
        }

    }
}
