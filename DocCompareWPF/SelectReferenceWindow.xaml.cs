using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;

namespace DocCompareWPF
{
    /// <summary>
    /// Interaktionslogik für SelectReferenceWindow.xaml
    /// </summary>
    public partial class SelectReferenceWindow : Window
    {
        public int selectedIndex;

        public SelectReferenceWindow()
        {
            InitializeComponent();
        }

        public void Setup(List<string> documentNames)
        {
            ObservableCollection<string> items = new ObservableCollection<string>();
            foreach (string name in documentNames)
            {
                items.Add(name);
            }

            fileNameComboBox.ItemsSource = items;
            fileNameComboBox.SelectedIndex = 0;
        }

        private void OkayButton_Click(object sender, RoutedEventArgs e)
        {
            selectedIndex = fileNameComboBox.SelectedIndex;
            DialogResult = true;
        }
    }
}
