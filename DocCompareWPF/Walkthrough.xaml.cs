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
            item.Text = "Thank you for choosing 2|Compare. In this quick guide, you will learn how to navigate and operate the software. Please click the NEXT button below to start the guide.";
            item.PathToFile = Path.Join(WalkthroughImageDirectory, "LogoLarge.png");
        }

        private void SkipButton_Click(object sender, RoutedEventArgs e)
        {
            CustomMessageBox msgBox = new CustomMessageBox();
            msgBox.Setup("Skip walkthrough", "Skip the entire walkthrough? You can restart the walkthrough from the settingspage.", "No", "Yes");

            if (msgBox.ShowDialog() == true) // skip walkthrough
            {
                DialogResult = true;
            }
            else
            {
                // do nothing
            }
        }

        private void SetText()
        {
            NextStepButton.Content = "Next";

            switch (stepCounter)
            {
                case 0:
                    item.Text = "Thank you for choosing 2|Compare. In this quick guide, you will learn how to navigate and operate the software. Please click the NEXT button below to start the guide.";
                    break;
                case 1:
                    item.Text = "The leftmost tab contains the document browse and inspection functionality. It is always opened on startup, and you can always return here by clicking this documents icon.";
                    break;
                case 2:
                    item.Text = "Click this button to browse for files. Alternatively, you can drag and drop up to five files onto the white canvas for subsequent document comparison.";
                    break;
                case 3:
                    item.Text = "The dropdown menus allow you to select loaded documents as the reference. The loaded documents will be automatically aligned to the selected reference document.";
                    break;
                case 4:
                    item.Text = "This button allows you to inspect document attributes, such as recent change dates or authors.";
                    break;
                case 5:
                    item.Text = "After document selection, click on this button to compare the selected document with the reference.";
                    break;
                case 6:
                    item.Text = "Document differences are highlighted in magenta. The highlights can be toggled on or off using this button.";
                    break;
                case 7:
                    item.Text = "Our algorithm aligns the pages of your documents automatically. If you would like to change the alignment manually, mouse over and click the link-icon on any page in the miniature preview. Then, select any page in the second document preview. The software automatically calculates a new optimal document alignment that incorporates your manual choice.";
                    break;
                case 8:
                    item.Text = "Click here to remove all manually set alignments. Click on the link icon between pages to remove individual links.";
                    break;
                case 9:
                    item.Text = "If you wish to amend files quickly, click this button to open the file in your external editor.";
                    break;
                case 10:
                    item.Text = "After changing a document externally, please reload the file using this button.";
                    break;
                case 11:
                    item.Text = "Use the dropdown menu to easily select other loaded document versions.";
                    break;
                case 12:
                    item.Text = "When differences are difficult to spot, mouse over the page, then click and hold down this icon. It will overlay the aligned pages for better visual comparison.";
                    break;
                case 13:
                    item.Text = "If the loaded Powerpoint slide contains a speaker note, click on this icon to inspect the changes.";
                    break;
                case 14:
                    item.Text = "The changes in speaker notes are highlighted in magenta.";
                    break;
                case 15:
                    item.Text = "If you have obtained a license key from www.hopie.tech, please activate your software under the settings tab.";
                    break;
                case 16:
                    item.Text = "Thanks for taking the time to complete the walkthrough. You can restart it any time under the settings tab. Enjoy 2|Compare!";
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

                if (stepCounter == 0)
                    item.PathToFile = Path.Join(WalkthroughImageDirectory, "LogoLarge.png");
                else
                    item.PathToFile = Path.Join(WalkthroughImageDirectory, "Page" + stepCounter.ToString() + ".PNG");
            }

            if (stepCounter == 0)
            {
                PreviousStepButton.IsEnabled = false;
            }
        }

        private void NextStepButton_Click(object sender, RoutedEventArgs e)
        {
            if (stepCounter <= 16)
            {
                stepCounter++;

                if (stepCounter == 17)
                {
                    DialogResult = true;
                }
                else
                {
                    SetText();
                    item.PathToFile = Path.Join(WalkthroughImageDirectory, "Page" + stepCounter.ToString() + ".PNG");
                }
            }

            if (stepCounter > 0)
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
