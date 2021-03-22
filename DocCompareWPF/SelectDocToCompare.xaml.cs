using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace DocCompareWPF
{
    /// <summary>
    /// Interaktionslogik für SelectDocToCompare.xaml
    /// </summary>
    public partial class SelectDocToCompare : Window
    {
        public int selectedIndex;

        public SelectDocToCompare()
        {
            InitializeComponent();
        }

        public void Setup(List<string> documentNames)
        {
            ObservableCollection<string> items = new ObservableCollection<string> ();
            foreach(string name in documentNames)
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
