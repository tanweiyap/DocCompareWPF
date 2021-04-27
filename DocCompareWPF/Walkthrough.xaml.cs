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
                    item.Text = "This dropdown menu allows you to select one loaded document as reference. All other documents will automatically be aligned to this document.";
                    break;
                case 4:
                    item.Text = "This button allows you to inspect document attributes, such as recent change dates or authors.";
                    break;
                case 5:
                    item.Text = "Click this button to enable the magnifying glass for inspecting small details of the document. The magnify function can also be activated by holding down the LEFT-CTRL key.";
                    break;
                case 6:
                    item.Text = "After reference selection, click this button at the top of any document to enter detailed comparison to the reference.";
                    break;
                case 7:
                    item.Text = "Document differences are highlighted in magenta. The highlights can be toggled on or off using this button.";
                    break;
                case 8:
                    item.Text = "Our algorithm aligns the pages of your documents automatically. If you would like to change the alignment manually, mouse over and click the link-icon on any page in the miniature preview. Then, select any page in the second document preview to which you want to align the previously selected page. The software calculates a new alignment with respect to your manual choice.";
                    break;
                case 9:
                    item.Text = "Click here to remove all manually set alignments. Click on an individual pages’ link icon to remove only that link.";
                    break;
                case 10:
                    item.Text = "If you wish to amend files quickly, click this button to open the file in your external editor.";
                    break;
                case 11:
                    item.Text = "After changing a document externally, please reload the file using this button.";
                    break;
                case 12:
                    item.Text = "Use the dropdown menu to easily select other loaded document versions.";
                    break;
                case 13:
                    item.Text = "When differences are difficult to spot, mouse over the page, then click and hold down this icon. It will overlay the aligned pages for better visual comparison.";
                    break;
                case 14:
                    item.Text = "If a PowerPoint slide contains speaker notes, you can access them through this icon. If notes contain differences, the Icon will be highlighted in Magenta.";
                    break;
                case 15:
                    item.Text = "Differences in speaker notes are highlighted as well. ";
                    break;
                case 16:
                    item.Text = "If you have obtained a license key from www.hopie.tech, please activate your software under the settings tab.";
                    break;
                case 17:
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
            if (stepCounter <= 17)
            {
                stepCounter++;

                if (stepCounter == 18)
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
