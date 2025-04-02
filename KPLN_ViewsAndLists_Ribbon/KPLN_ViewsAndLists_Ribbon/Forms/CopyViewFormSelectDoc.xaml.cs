using System;
using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class DocumentSelectionWindow : Window
    {
        private readonly Dictionary<string, Document> documentDictionary = new Dictionary<string, Document>();
        public Document SelectedDocument { get; private set; }

        public DocumentSelectionWindow(List<Document> documents)
        {
            InitializeComponent();

            foreach (var doc in documents)
            {
                documentDictionary[doc.Title] = doc;
                DocumentListBox.Items.Add(doc.Title);
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            if (DocumentListBox.SelectedItem is string selectedTitle && documentDictionary.ContainsKey(selectedTitle))
            {
                SelectedDocument = documentDictionary[selectedTitle];
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("Выберите документ.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
