using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;

namespace DocCompareWPF
{
    /// <summary>
    /// Interaktionslogik für Walkthrough.xaml
    /// </summary>
    public partial class Walkthrough : Window
    {

        string WalkthroughImageDirectory;
        WalkthroughImageItem item;
        int stepCounter = 0;
        public Walkthrough()
        {
            InitializeComponent();
            item = new WalkthroughImageItem();
            
            WalkthroughImageDirectory = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
            WalkthroughImageDirectory = Path.Join(WalkthroughImageDirectory, "WalkthroughImages");

            MessageTextBlock.DataContext = item;
            ImageView.DataContext = item;
            InitMessage();
        }

        private void InitMessage()
        {
            item.Text = "Thanks for choosing 2|Compare. We will now guide you through the application to get you familiar with the functionality.";
            item.PathToFile = Path.Join(WalkthroughImageDirectory, "LogoLarge.png");
        }

        private void SkipButton_Click(object sender, RoutedEventArgs e)
        {
            CustomMessageBox msgBox = new CustomMessageBox();
            msgBox.Setup("Skip walkthrough", "Skip the entire walkthrough? You can restart the walkthrough from the settingspage.", "No", "Yes");

            if(msgBox.ShowDialog() == true) // skip walkthrough
            {
                DialogResult = true;
            }else
            {
                // do nothing
            }            
        }

        private void SetText()
        {
            switch (stepCounter)
            {
                case 0:
                    item.Text = "Thanks for choosing 2|Compare. We will now guide you through the application to get you familiar with the functionality.";
                    break;
                case 1:
                    item.Text = "Click on this tab to browse and inspect files before comparison.";
                    break;
                case 2:
                    item.Text = "Click on this button to browse files for inspection. Alternatively, you can drag and drop files into this zone to load them. A total of 5 files can be loaded at once.";
                    break;
                case 3:
                    item.Text = "Use the dropdown menu to view other loaded documents. The selected document on the left will be the reference document for the comparison.";
                    break;
                case 4:
                    item.Text = "Click on the info button to view more document attributes.";
                    break;
                case 5:
                    item.Text = "Click on this tab to switch to compare mode after document selection.";
                    break;
                case 6:
                    item.Text = "Document differences are highlighted. Click on this button to toggle highlighting.";
                    break;
                case 7:
                    item.Text = "Pages are aligned automatically. If you would like to change the alignment manually, mouse over and click the 'link' icon on a page in the miniature preview, then choose any page in the second document to align these two pages.";
                    break;
                case 8:
                    item.Text = "Click here to remove all previously set link. Click on the icon between pages to remove individual links.";
                    break;
                case 9:
                    item.Text = "Click this button to open the file in external editor.";
                    break;
                case 10:
                    item.Text = "Click this button to reload the file after editing for comparison.";
                    break;
                case 11:
                    item.Text = "Use the dropdown menu to select other document for comparison with the reference document.";
                    break;
                case 12:
                    item.Text = "If you have made a license subscription on www.hopie.tech, you can activate the license under the settings menu.";
                    break;
                case 13:
                    item.Text = "Thanks for completing the walkthrough guide. Enjoy 2|Compare! You can restart this walkthrough from the settings page.";
                    NextStepButton.Content = "Close";
                    break;
            }
        }

        private void PreviousStepButton_Click(object sender, RoutedEventArgs e)
        {
            if (stepCounter != 0)
            {
                stepCounter--;

                SetText();

                if(stepCounter == 0)
                    item.PathToFile = Path.Join(WalkthroughImageDirectory, "LogoLarge.png");
                else
                    item.PathToFile = Path.Join(WalkthroughImageDirectory, "Page" + stepCounter.ToString() + ".PNG");
            }

            if(stepCounter == 0)
            {
                PreviousStepButton.IsEnabled = false;
            }
        }

        private void NextStepButton_Click(object sender, RoutedEventArgs e)
        {
            if (stepCounter <= 13)
            {
                stepCounter++;

                if (stepCounter == 14)
                {
                    DialogResult = true;
                }
                else
                {
                    SetText();
                    item.PathToFile = Path.Join(WalkthroughImageDirectory, "Page" + stepCounter.ToString() + ".PNG");
                }
            }

            if(stepCounter > 0)
            {
                PreviousStepButton.IsEnabled = true;
            }
        }
    }

    public class WalkthroughImageItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _pathToFile;
        public string PathToFile
        {
            get
            {
                return _pathToFile;
            }

            set
            {
                _pathToFile = value;
                OnPropertyChanged();
            }
        }

        private string _text;
        public string Text
        {
            get
            {
                return _text;
            }

            set
            {
                _text = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
