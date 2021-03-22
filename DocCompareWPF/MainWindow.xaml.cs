using DocCompareWPF.Classes;
using Microsoft.Win32;
using ProtoBuf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Xml;

namespace DocCompareWPF
{
    public static class TextBlockHelper
    {
        #region FormattedText Attached dependency property

        public static string GetFormattedText(DependencyObject obj)
        {
            return (string)obj.GetValue(FormattedTextProperty);
        }

        public static void SetFormattedText(DependencyObject obj, string value)
        {
            obj.SetValue(FormattedTextProperty, value);
        }

        public static readonly DependencyProperty FormattedTextProperty =
            DependencyProperty.RegisterAttached("FormattedText",
            typeof(string),
            typeof(TextBlockHelper),
            new UIPropertyMetadata("", FormattedTextChanged));

        private static void FormattedTextChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            string value = e.NewValue as string;

            TextBlock textBlock = sender as TextBlock;

            if (textBlock != null)
            {
                textBlock.Inlines.Clear();
                textBlock.Inlines.Add(Process(value));
            }
        }

        #endregion

        static Inline Process(string value)
        {
            XmlDocument doc = new XmlDocument();

            if (value != null)
                doc.LoadXml(value);

            Span span = new Span();

            if (doc.ChildNodes.Count != 0)
                InternalProcess(span, doc.ChildNodes[1]);

            return span;
        }

        private static void InternalProcess(Span span, XmlNode xmlNode)
        {
            foreach (XmlNode child in xmlNode)
            {
                if (child is XmlText)
                {
                    span.Inlines.Add(new Run(child.InnerText));
                }
                else if (child is XmlElement)
                {
                    Span spanItem = new Span();
                    InternalProcess(spanItem, child);
                    switch (child.Name.ToUpper())
                    {
                        case "B":
                        case "BOLD":
                            Bold bold = new Bold(spanItem);
                            span.Inlines.Add(bold);
                            break;
                        case "I":
                        case "ITALIC":
                            Italic italic = new Italic(spanItem);
                            span.Inlines.Add(italic);
                            break;
                        case "U":
                        case "UNDERLINE":
                            Underline underline = new Underline(spanItem);
                            span.Inlines.Add(underline);
                            break;
                        case "D":
                        case "DELETE":
                            spanItem.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                            spanItem.Foreground = Brushes.Transparent;
                            span.Inlines.Add(spanItem);
                            break;
                        case "IN":
                        case "INSERT":
                            spanItem.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                            span.Inlines.Add(spanItem);
                            break;
                    }
                }
            }
        }
    }

    public class CompareMainItem : INotifyPropertyChanged
    {
        private Visibility _showMask;

        public string Document1 { get; set; }
        public string Document2 { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public bool AniDiffButtonEnable { get; set; }
        public string AnimateDiffLeftButtonName { get; set; }
        public string AnimateDiffRightButtonName { get; set; }
        public string ImgAniLeftName { get; set; }
        public string ImgAniRightName { get; set; }
        public string ImgGridLeftName { get; set; }
        public string ImgGridName { get; set; }
        public string ImgGridRightName { get; set; }
        public string ImgLeftName { get; set; }
        public string ImgMaskRightName { get; set; }
        public string ImgRightName { get; set; }
        public Thickness Margin { get; set; }

        private string _pathToAniImgLeft;
        private string _pathToAniImgRight;
        private string _pathToImgLeft;
        private string _pathToImgRight;
        private string _pathToMaskImgRight;
        public string PathToAniImgLeft { get { return _pathToAniImgLeft; } set { _pathToAniImgLeft = value; OnPropertyChanged(); } }
        public string PathToAniImgRight { get { return _pathToAniImgRight; } set { _pathToAniImgRight = value; OnPropertyChanged(); } }
        public string PathToImgLeft { get { return _pathToImgLeft; } set { _pathToImgLeft = value; OnPropertyChanged(); } }
        public string PathToImgRight { get { return _pathToImgRight; } set { _pathToImgRight = value; OnPropertyChanged(); } }
        public string PathToMaskImgRight { get { return _pathToMaskImgRight; } set { _pathToMaskImgRight = value; OnPropertyChanged(); } }

        private double _blurRadiusLeft;
        private double _blurRadiusRight;

        private Visibility _showHiddenLeft;
        private Visibility _showHiddenRight;
        private Visibility _pptNoteGridLeftVisi;
        private Visibility _pptNoteGridRightVisi;

        public string PPTSpeakerNoteGridNameLeft { get; set; }
        public string PPTSpeakerNoteGridNameRight { get; set; }
        public string ClosePPTSpeakerNotesButtonNameLeft { get; set; }
        public string ClosePPTSpeakerNotesButtonNameRight { get; set; }
        //public string PPTSpeakerNotesLeft { get; set; }
        //public string PPTSpeakerNotesRight { get; set; }
        public string ShowPPTSpeakerNotesButtonNameLeft { get; set; }
        public string ShowPPTSpeakerNotesButtonNameRight { get; set; }
        public string ShowPPTSpeakerNotesButtonNameRightChanged { get; set; }
        private Visibility _showPPTSpeakerNotesButtonLeft;
        private Visibility _showPPTSpeakerNotesButtonRight;
        private Visibility _showPPTSpeakerNotesButtonRightChanged;

        private SolidColorBrush _showPPTButtonBackground;
        public SolidColorBrush ShowPPTNoteButtonBackground
        {
            get
            {
                return _showPPTButtonBackground;
            }

            set
            {
                _showPPTButtonBackground = value;
                OnPropertyChanged();
            }
        }

        public Visibility showPPTSpeakerNotesButtonLeft
        {
            get
            {
                return _showPPTSpeakerNotesButtonLeft;
            }

            set
            {
                _showPPTSpeakerNotesButtonLeft = value;
                OnPropertyChanged();
            }
        }

        public Visibility showPPTSpeakerNotesButtonRight
        {
            get
            {
                return _showPPTSpeakerNotesButtonRight;
            }

            set
            {
                _showPPTSpeakerNotesButtonRight = value;
                OnPropertyChanged();
            }
        }
        public Visibility showPPTSpeakerNotesButtonRightChanged
        {
            get
            {
                return _showPPTSpeakerNotesButtonRightChanged;
            }

            set
            {
                _showPPTSpeakerNotesButtonRightChanged = value;
                OnPropertyChanged();
            }
        }

        public Visibility PPTNoteGridLeftVisi
        {
            get
            {
                return _pptNoteGridLeftVisi;
            }

            set
            {
                _pptNoteGridLeftVisi = value;
                OnPropertyChanged();
            }
        }
        public Visibility PPTNoteGridRightVisi
        {
            get
            {
                return _pptNoteGridRightVisi;
            }

            set
            {
                _pptNoteGridRightVisi = value;
                OnPropertyChanged();
            }
        }

        public double BlurRadiusRight
        {
            get
            {
                return _blurRadiusRight;
            }

            set
            {
                _blurRadiusRight = value;
                OnPropertyChanged();
            }
        }
        public double BlurRadiusLeft
        {
            get
            {
                return _blurRadiusLeft;
            }

            set
            {
                _blurRadiusLeft = value;
                OnPropertyChanged();
            }
        }



        public Visibility ShowMask
        {
            get
            {
                return _showMask;
            }

            set
            {
                _showMask = value;
                OnPropertyChanged();
            }
        }

        public Visibility ShowHiddenLeft
        {
            get
            {
                return _showHiddenLeft;
            }

            set
            {
                _showHiddenLeft = value;
                OnPropertyChanged();
            }
        }
        public Visibility ShowHiddenRight
        {
            get
            {
                return _showHiddenRight;
            }

            set
            {
                _showHiddenRight = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string appDataDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare");
        private readonly DocumentManagement docs;
        private readonly string versionString = "1.1.1";
        private readonly string localetype = "DE";
        private readonly string workingDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare");
        private string compareResultFolder;
        private bool docCompareRunning, docProcessRunning, animateDiffRunning, showMask;
        private int docCompareSideGridShown, docProcessingCounter;
        private Grid gridToAnimate;
        private bool inForceAlignMode;
        private string lastUsedDirectory;

        // License management
        private LicenseManagement lic;

        private string licKeyLastInputString;
        private string trialKeyLastInputString;

        // Update
        private string updateInstallerURL;

        // Stack panel for viewing documents in scrollviewer control in comparison view
        //private VirtualizingStackPanel childPanel1;
        private StackPanel refDocPanel;

        private double scrollPosLeft, scrollPosRight;
        private string selectedSideGridButtonName1 = "";
        private string selectedSideGridButtonName2 = "";

        // App settings
        private AppSettings settings;

        private GridSelection sideGridSelectedLeftOrRight, mainGridSelectedLeftOrRight;
        private Thread threadLoadDocs;
        private Thread threadLoadDocsProgress;
        private Thread threadCompare;
        private Thread threadAnimateDiff;
        private Thread threadDisplayResult;
        private Thread threadCheckTrial;
        private Thread threadCheckUpdate;
        private readonly Thread threadRenewLic;
        private Thread threadStartWalkthrough;

        // Walkthrough
        private bool walkthroughMode;

        // link scroll
        private bool linkscroll = true;

        private WalkthroughSteps walkthroughStep = 0;

        // mouse over hidden ppt slides effect buffer
        Effect hiddenPPTEffect;
        Visibility hiddenPPTVisi;

        SidePanels currentVisiblePanel;

        public MainWindow()
        {
            InitializeComponent();
            Directory.CreateDirectory(appDataDir);
            compareResultFolder = Path.Join(workingDir, Guid.NewGuid().ToString());

            // GUI stuff
            showMask = true;
            AppVersionLabel.Content = "Version " + versionString;
            SetVisiblePanel(SidePanels.DRAGDROP);
            SidePanelDocCompareButton.IsEnabled = false;
            ActivateLicenseButton.IsEnabled = false;

            HideDragDropZone2();
            HideDragDropZone3();
            HideDragDropZone4();
            HideDragDropZone5();

            try
            {
                LoadSettings();
                lastUsedDirectory = settings.defaultFolder;

                //settings.isProVersion = true;
                //settings.canSelectRefDoc = true;
            }
            catch
            {
                settings = new AppSettings
                {
                    defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    numPanelsDragDrop = 2
                };

                SaveSettings();
            }

            docs = new DocumentManagement(settings.maxDocCount, workingDir, settings);

            // License Management
            try
            {
                LoadLicense();
                DisplayLicense();

                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
                {
                    threadCheckTrial = new Thread(new ThreadStart(CheckTrial));
                    threadCheckTrial.Start();
                }

                ErrorHandling.ReportStatus("App Launch", "App version: " + versionString + " successfully launched with license: " + lic.GetLicenseTypesString() + ", expires/renewal on " + lic.GetExpiryDateString() + " on " + lic.GetUUID());
            }
            catch
            {
                lic = new LicenseManagement();
                lic.Init(); // init 7 days trial
                DisplayLicense();
                SaveLicense();

                CustomMessageBox msgBox = new CustomMessageBox();
                msgBox.Setup("Product activation", "A trial license for 7 days has been started. If you have subscribed to a license, proceed to activating your subcription under the settings menu.", "Okay");
                msgBox.ShowDialog();

                threadCheckTrial = new Thread(new ThreadStart(CheckTrial));
                threadCheckTrial.Start();

                ErrorHandling.ReportStatus("New trial license", "on " + lic.GetUUID() + ", Expires on " + lic.GetExpiryDateString());
            }

            TimeSpan timeBuffer = lic.GetExpiryDate().Subtract(DateTime.Today);

            // Reminder to subscribe
            if (timeBuffer.TotalDays <= 5 && timeBuffer.TotalDays > 0 && lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
            {
                CustomMessageBox msgBox = new CustomMessageBox();
                msgBox.Setup("Expired lincense", "Your trial license will expire in " + timeBuffer.TotalDays + " day(s). Please consider making a subscription on www.hopietech.com", "Okay");
                msgBox.ShowDialog();
            }

            // if license expires or needs renewal
            if (timeBuffer.TotalDays <= 0)
            {
                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
                {
                    if (settings.trialExtended == false)
                    {
                        CustomMessageBox msgBox = new CustomMessageBox();
                        msgBox.Setup("Expired lincense", "Your license has expired. If you wish to extend the trial for 7 days, feel free to fill up a feedback survey on our website.", "Cancel", "Take survey");

                        if (msgBox.ShowDialog() == true) // user wish to take survey
                        {
                            ProcessStartInfo info = new ProcessStartInfo("https://de.hopie.tech/quiz/fragen-zur-verlaengerung-der-testphase/")
                            {
                                UseShellExecute = true
                            };
                            Process.Start(info);

                            settings.showExtendTrial = true;
                            SaveSettings();
                        }
                    }

                    DisplayLicense();
                    BrowseFileButton1.IsEnabled = false;
                    DocCompareFirstDocZone.AllowDrop = false;
                    DocCompareDragDropZone1.AllowDrop = false;
                    DocCompareColorZone1.AllowDrop = false;
                }

                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.ANNUAL_SUBSCRIPTION)
                {
                    threadRenewLic = new Thread(new ThreadStart(RenewLicense));
                    threadRenewLic.Start();
                }
            }

            // Extend trial option
            if (settings.showExtendTrial == true)
            {
                ExtendTrialGrid1.Visibility = Visibility.Visible;
                ExtendTrialGrid2.Visibility = Visibility.Visible;
            }

            // Walkthrough if needed
            threadStartWalkthrough = new Thread(new ThreadStart(ShowWalkthroughStartMessage));
            threadStartWalkthrough.Start();

            // Check update if needed
            threadCheckUpdate = new Thread(new ThreadStart(CheckUpdate));
            threadCheckUpdate.Start();
        }

        private enum GridSelection
        {
            LEFT,
            RIGHT,
        };

        private enum SettingsPanels
        {
            DOCBROWSING,
            DOCCOMPARE,
            SUBSCRIPTION,
            ABOUT,
        };

        private enum SidePanels
        {
            DRAGDROP,
            DOCCOMPARE,
            REFDOC,
            FILE_EXPLORER,
            SETTINGS,
        };

        private enum WalkthroughSteps
        {
            START,
            BROWSEFILETAB,
            BROWSEFILEBUTTON1,
            BROWSEFILEBUTTON2,
            BROWSEFILECOMBOBOX,
            BROWSEFILEOPENFURTHERFILES,
            BROWSEFILECLOSEFILES,
            BROWSEFILEINFOOPEN,
            BROWSEFILEINFOCLOSE,
            COMPARETAB,
            COMPAREHIGHLIGHT,
            COMPAREOPENEXTERN,
            COMPARERELOAD,
            COMPARELINK,
            COMPAREUNLINK,
            COMPAREANIMATE,
            END,
        }

        private void ShowWalkthroughStartMessage()
        {
            try
            {
                Thread.Sleep(500);

                // First time launch app?
                if (settings.shownWalkthrough == false)
                {
                    Dispatcher.Invoke(() =>
                    {
                        Walkthrough walkthrough = new Walkthrough();
                        if (walkthrough.ShowDialog() == true)
                        {
                            settings.shownWalkthrough = true;
                            SaveSettings();
                        }

                        /*
                        CustomMessageBox msgBox = new CustomMessageBox();
                        msgBox.Setup("Application walkthrough", "Thanks for choosing 2|Compare. We will now guide you through the application to get you familiar with the functionality", "Proceed", "Skip");

                        if (msgBox.ShowDialog() == true) // user wish to skip walkthrough
                        {
                            settings.shownWalkthrough = true;
                            SaveSettings();
                        }
                        else
                        {
                            walkthroughMode = true;
                            walkthroughStep = WalkthroughSteps.BROWSEFILETAB;
                            PopupBrowseFileBubble.IsOpen = true;
                        }
                        */
                    });
                }
            }
            catch
            {
                // thread aborted. maybe the user has close the program before we show start the walkthrough
            }
        }

        private async void ActivateLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            // remove anyway
            bool res_1 = await lic.RemoveLicense();

            LicenseManagement.LicServerResponse res = await lic.ActivateLicense(UserEmailTextBox.Text, LicenseKeyTextBox.Text);
            CustomMessageBox msgBox;
            switch (res)
            {
                case LicenseManagement.LicServerResponse.UNREACHABLE:
                    msgBox = new CustomMessageBox();
                    msgBox.Setup("No connection to license server", "The 2|Compare license server cannot be reached. This may be due to multiple reasons, including your firewall settings or no connection to the Internet but also server maintenance on our side. If you are connected to the Internet and your firewall is configured properly, please ignore this warning. Out server will be up and running again shortly. If you see this warning throughout multiple consecutive days, please contact us at support@hopie.tech.", "Okay");
                    msgBox.ShowDialog();
                    break;

                case LicenseManagement.LicServerResponse.KEY_MISMATCH:
                    msgBox = new CustomMessageBox();
                    msgBox.Setup("Invalid license key", "The provided license key does not match the email address. Please check your inputs.", "Okay");
                    msgBox.ShowDialog();
                    break;

                case LicenseManagement.LicServerResponse.ACCOUNT_NOT_FOUND:
                    msgBox = new CustomMessageBox();
                    msgBox.Setup("License not found", "No license was found under the given email address. Please check your inputs.", "Okay");
                    msgBox.ShowDialog();
                    break;

                case LicenseManagement.LicServerResponse.OKAY:
                    msgBox = new CustomMessageBox();
                    msgBox.Setup("License activation", "License activated successfully.", "Okay");
                    msgBox.ShowDialog();
                    UserEmailTextBox.IsEnabled = false;
                    LicenseKeyTextBox.IsEnabled = false; // after successful activation, we will prevent further editing
                    ActivateLicenseButton.IsEnabled = false;
                    SaveLicense(); // only save license info if successful
                    // allow usage if it was previously disabled
                    BrowseFileButton1.IsEnabled = true;
                    DocCompareFirstDocZone.AllowDrop = true;
                    DocCompareDragDropZone1.AllowDrop = true;
                    DocCompareColorZone1.AllowDrop = true;
                    break;

                case LicenseManagement.LicServerResponse.INVALID:
                    msgBox = new CustomMessageBox();
                    msgBox.Setup("License activation failed", "License activated failed. Please try again later", "Okay");
                    msgBox.ShowDialog();
                    // we do nothing and retain current license info
                    break;

                case LicenseManagement.LicServerResponse.INUSE:
                    msgBox = new CustomMessageBox();
                    msgBox.Setup("License in use", "License has been activated on another machine. Please contact support@hopietech.com for further assitance.", "Okay");
                    msgBox.ShowDialog();
                    // we do nothing and retain current license info
                    break;
            }

            DisplayLicense();
        }

        private void AnimateDiffThread()
        {
            bool imageToggler = true;

            while (animateDiffRunning)
            {
                Dispatcher.Invoke(() =>
                {
                    // Turn off Mask and animate
                    foreach (object child in gridToAnimate.Children)
                    {
                        if (child is Border)
                        {
                            Image thisImg = (child as Border).Child as Image;
                            if (thisImg.Tag.ToString().Contains("Ani"))
                            {
                                if (imageToggler == false)
                                    thisImg.Visibility = Visibility.Hidden;
                                else
                                    thisImg.Visibility = Visibility.Visible;
                            }
                            else if (thisImg.Tag.ToString().Contains("Mask"))
                            {
                            }
                            else
                            {
                                if (imageToggler == false)
                                    thisImg.Visibility = Visibility.Visible;
                                else
                                    thisImg.Visibility = Visibility.Hidden;
                            }
                        }

                        if (child is Grid)
                        {
                            Grid thisGrid = child as Grid;
                            thisGrid.Visibility = Visibility.Hidden;
                        }
                    }

                    if (imageToggler == false)
                    {
                        imageToggler = true;
                    }
                    else
                    {
                        imageToggler = false;
                    }
                });

                Thread.Sleep(400);
            }
        }

        private void BrowseFileButton1_Click(object sender, RoutedEventArgs e)
        {
            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON1)
            {
                PopupBrowseFileButtonBubble.IsOpen = false;
                lastUsedDirectory = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
                lastUsedDirectory = Path.Join(lastUsedDirectory, "examples");
            }

            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif, *bmp)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF;*.bmp;*.BMP|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF,*.bmp,*.BMP |All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                if (walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON1)
                {
                    walkthroughStep = WalkthroughSteps.BROWSEFILEBUTTON2;
                }

                string[] filenames = openFileDialog.FileNames;
                string ext;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);

                foreach (string file in filenames)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    SelectReferenceWindow selectReferenceWindow = new SelectReferenceWindow();
                    List<string> fileList = new List<string>();
                    foreach (Document doc in docs.documents)
                    {
                        fileList.Add(Path.GetFileName(doc.filePath));
                    }

                    selectReferenceWindow.Setup(fileList);

                    if (docs.documents.Count >= 2)
                    {
                        if (selectReferenceWindow.ShowDialog() == true)
                        {
                            int desiredInd = selectReferenceWindow.selectedIndex;

                            int existingInd = -1;
                            for (int i = 0; i < docs.documents.Count; i++)
                            {
                                if (docs.documentsToShow[i] == desiredInd)
                                {
                                    existingInd = i;
                                }
                            }

                            docs.documentsToShow[existingInd] = docs.documentsToShow[0];
                            docs.documentsToShow[0] = selectReferenceWindow.selectedIndex;
                        }
                    }

                    /*
                    if ((sender as Button).Name.Contains("Top"))
                    {
                        if (docs.documents.Count >= 3)
                        {
                            docs.documentsToShow[0] = docs.documents.Count - 1;
                        }
                        else if (docs.documents.Count == 2)
                        {
                            docs.documentsToShow[1] = 1;
                        }
                    }
                    */
                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
            else
            {
                if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON1)
                {
                    PopupBrowseFileButtonBubble.IsOpen = true;
                }
            }
        }

        private void BrowseFileButton2_Click(object sender, RoutedEventArgs e)
        {
            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON2)
            {
                PopupBrowseFileButton2Bubble.IsOpen = false;
                lastUsedDirectory = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
                lastUsedDirectory = Path.Join(lastUsedDirectory, "examples");
            }

            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif, *bmp)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF;*.bmp;*.BMP|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF,*.bmp,*.BMP |All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filenames = openFileDialog.FileNames;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);
                string ext;

                foreach (string file in filenames)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    /*
                    if ((sender as Button).Name.Contains("Top"))
                    {
                        if (docs.documents.Count >= 2)
                        {
                            docs.documentsToShow[1] = docs.documents.Count - 1;
                        }
                    }
                    */

                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
            else
            {
                if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON2)
                {
                    PopupBrowseFileButton2Bubble.IsOpen = true;
                }
            }
        }

        private void BrowseFileButton3_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif, *bmp)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF;*.bmp;*.BMP|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF,*.bmp,*.BMP |All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filenames = openFileDialog.FileNames;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);
                string ext;

                foreach (string file in filenames)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }

        private async void CheckTrial()
        {
            LicenseManagement.LicServerResponse res = await lic.ActivateTrial();
            if (res == LicenseManagement.LicServerResponse.INVALID)
            {
                Dispatcher.Invoke(() =>
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("Expired lincense", "An expired trial license was detected for your computer. If you wish to continue to use 2|Compare, please purchase a license on https://hopie.tech.", "Okay");
                    msgBox.ShowDialog();

                    BrowseFileButton1.IsEnabled = false;
                    DocCompareFirstDocZone.AllowDrop = false;
                    DocCompareDragDropZone1.AllowDrop = false;
                    DocCompareColorZone1.AllowDrop = false;
                });
            }

            Dispatcher.Invoke(() => { DisplayLicense(); });
        }

        private async void CheckUpdate()
        {
            Thread.Sleep(7000);
            try
            {
                List<string> res = await lic.CheckUpdate(versionString, localetype);
                if (res != null)
                {
                    updateInstallerURL = res[1];
                    Dispatcher.Invoke(() =>
                    {
                        WindowUpdateButton.Visibility = Visibility.Visible;

                        if (res[0] != settings.skipVersionString)
                        {
                            CustomMessageBox msgBox = new CustomMessageBox();
                            msgBox.Setup("Update available", "A newer version of 2|Compare is available. Click OKAY to proceed with downloading the installer.", "Okay", "Skip");

                            if (msgBox.ShowDialog() == true)
                            {
                                settings.skipVersion = true;
                                settings.skipVersionString = res[0];
                                SaveSettings();
                            }
                            else
                            {
                                updateInstallerURL = res[1];
                                ProcessStartInfo info = new ProcessStartInfo(updateInstallerURL)
                                {
                                    UseShellExecute = true
                                };
                                Process.Start(info);
                            }
                        }
                    });
                }
                else
                {
                    Dispatcher.Invoke(() =>
                    {
                        WindowUpdateButton.Visibility = Visibility.Hidden;
                    });
                }
            }catch
            {

            }
        }

        private async void ExtendTrial()
        {
            LicenseManagement.LicServerResponse res = await lic.ExtendTrial();
            if (res == LicenseManagement.LicServerResponse.INVALID)
            {
                Dispatcher.Invoke(() =>
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("Expired lincense", "An expired trial license was detected for your computer. If you wish to continue to use 2|Compare, please purchase a license on https://hopie.tech.", "Okay");
                    msgBox.ShowDialog();

                    BrowseFileButton1.IsEnabled = false;
                    DocCompareFirstDocZone.AllowDrop = false;
                    DocCompareDragDropZone1.AllowDrop = false;
                    DocCompareColorZone1.AllowDrop = false;

                    DisplayLicense();
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("Trial extended", "Your trial has been successfully extended", "Okay");
                    msgBox.ShowDialog();

                    BrowseFileButton1.IsEnabled = true;
                    DocCompareFirstDocZone.AllowDrop = true;
                    DocCompareDragDropZone1.AllowDrop = true;
                    DocCompareColorZone1.AllowDrop = true;

                    DisplayLicense();
                });

                SaveLicense();

                settings.trialExtended = true;
                SaveSettings();
            }
        }

        private void CloseDoc1Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[0], 0);
            CloseDocumentCommonPart();
            UpdateDocSelectionComboBox();

            if (docs.documents.Count < 2)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void CloseDoc2Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[1], 1);
            CloseDocumentCommonPart();
            UpdateDocSelectionComboBox();

            if (docs.documents.Count < 2)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void CloseDoc3Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[2], 2);
            CloseDocumentCommonPart();
            UpdateDocSelectionComboBox();

            if (docs.documents.Count < 2)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void CloseDocumentCommonPart()
        {
            //TODO: Premium

            // hide all then show them individually
            HideDragDropZone2();
            HideDragDropZone3();
            HideDragDropZone4();
            HideDragDropZone5();

            if (docs.documents.Count == 0)
            {
                Doc1Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone1.Visibility = Visibility.Visible;
                DocPreviewStatGrid.Visibility = Visibility.Collapsed;
            }
            else
            {
                if (docs.documents.Count >= 1)
                {
                    ShowDragDropZone2();
                    DisplayPreview(1, docs.documentsToShow[0]);
                    Doc1StatsGrid.Visibility = Visibility.Visible;
                    Doc2Grid.Visibility = Visibility.Hidden;
                    Doc2StatsGrid.Visibility = Visibility.Collapsed;
                    Doc2NameLabel.Content = "";
                }

                if (docs.documents.Count >= 2)
                {
                    ShowDragDropZone3();
                    DisplayPreview(2, docs.documentsToShow[1]);
                    Doc2StatsGrid.Visibility = Visibility.Visible;
                    Doc3Grid.Visibility = Visibility.Hidden;
                    Doc3StatsGrid.Visibility = Visibility.Collapsed;
                    Doc3NameLabel.Content = "";
                }

                if (docs.documents.Count >= 3)
                {
                    ShowDragDropZone4();
                    DisplayPreview(3, docs.documentsToShow[2]);
                    Doc3StatsGrid.Visibility = Visibility.Visible;
                    Doc4Grid.Visibility = Visibility.Hidden;
                    Doc4StatsGrid.Visibility = Visibility.Collapsed;
                    Doc4NameLabel.Content = "";
                }

                if (docs.documents.Count >= 4)
                {
                    ShowDragDropZone5();
                    DisplayPreview(4, docs.documentsToShow[3]);
                    Doc4StatsGrid.Visibility = Visibility.Visible;
                    Doc5Grid.Visibility = Visibility.Hidden;
                    Doc5StatsGrid.Visibility = Visibility.Collapsed;
                    Doc5NameLabel.Content = "";
                }

                if (docs.documents.Count >= 5)
                {
                    DisplayPreview(5, docs.documentsToShow[4]);
                    Doc5StatsGrid.Visibility = Visibility.Visible;
                }

                ShowInfoButtonSetVisi();
            }
        }

        private void CompareDocsThread()
        {
            try
            {
                Dispatcher.Invoke(() => { UpdateFileStat(5); });

                docCompareRunning = true;

                // show mask should always be enabled
                showMask = true;
                Dispatcher.Invoke(() =>
                {
                    ShowMaskButton.Visibility = Visibility.Hidden;
                    HideMaskButton.Visibility = Visibility.Visible;
                    HighlightingDisableTip.Visibility = Visibility.Hidden;
                });

                int[,] forceIndices = new int[docs.forceAlignmentIndices.Count, 2];
                for (int i = 0; i < docs.forceAlignmentIndices.Count; i++)
                {
                    forceIndices[i, 0] = docs.forceAlignmentIndices[i][0];
                    forceIndices[i, 1] = docs.forceAlignmentIndices[i][1];
                }

                if (Directory.Exists(compareResultFolder))
                {
                    DirectoryInfo di = new DirectoryInfo(compareResultFolder);
                    di.Delete(true);
                }

                compareResultFolder = Path.Join(workingDir, Guid.NewGuid().ToString());
                Directory.CreateDirectory(compareResultFolder);

                Document.CompareDocs(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[1]].imageFolder, compareResultFolder, out docs.pageCompareIndices, out docs.totalLen, forceIndices);
                docs.documents[docs.documentsToCompare[0]].docCompareIndices = new List<int>();
                docs.documents[docs.documentsToCompare[1]].docCompareIndices = new List<int>();

                if (docs.totalLen != 0)
                {
                    for (int i = docs.totalLen - 1; i >= 0; i--)
                    {
                        docs.documents[docs.documentsToCompare[0]].docCompareIndices.Add((int)docs.pageCompareIndices[i]);
                        docs.documents[docs.documentsToCompare[1]].docCompareIndices.Add((int)docs.pageCompareIndices[i + docs.totalLen]);
                    }
                }

                docCompareRunning = false;

                // if is PPT, we will compare the speaker notes
                // we will only display the difference on the right
                docs.pptSpeakerNotesDiff = new List<List<DocCompareDLL.Diff>>(docs.totalLen);
                DocCompareDLL.diff_match_patch diffMatch = new DocCompareDLL.diff_match_patch();
                string text1 = "";
                string text2 = "";

                for (int i = 0; i < docs.totalLen; i++)
                {
                    if (docs.documents[docs.documentsToCompare[0]].fileType == Document.FileTypes.PPT &&
                    docs.documents[docs.documentsToCompare[1]].fileType == Document.FileTypes.PPT)
                    {
                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                        {
                            if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                            {
                                text1 = docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]];
                            }
                            else
                            {
                                text1 = "";
                            }
                        }

                        if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                        {
                            if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                            {
                                text2 = docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]];
                            }
                            else
                            {
                                text2 = "";
                            }
                        }

                        if (text1.Length != 0 && text2.Length != 0)
                        {
                            List<DocCompareDLL.Diff> diff = diffMatch.diff_main(text1, text2);
                            diffMatch.diff_cleanupSemantic(diff);
                            docs.pptSpeakerNotesDiff.Add(diff);
                        }
                        else
                        {
                            docs.pptSpeakerNotesDiff.Add(new List<DocCompareDLL.Diff>());
                        }
                    }
                    else
                    {
                        docs.pptSpeakerNotesDiff.Add(new List<DocCompareDLL.Diff>());
                    }
                }

                Dispatcher.Invoke(() =>
                {
                    threadDisplayResult = new Thread(new ThreadStart(DisplayComparisonResult));
                    ProgressBarLoadingResults.Visibility = Visibility.Visible;
                    ProgressBarDocCompare.Visibility = Visibility.Hidden;
                    threadDisplayResult.Start();
                    ShowDocCompareFileInfoButton.IsEnabled = true;

                    if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPAREHIGHLIGHT)
                    {
                        PopupHighlightOffBubble.IsOpen = true;
                        //walkthroughStep = WalkthroughSteps.COMPAREOPENEXTERN; // 8
                    }
                    else
                    {
                        PopupHighlightOffBubble.IsOpen = false;
                    }
                });
            }
            catch
            {
                docCompareRunning = false;
            }
        }

        private void DisableRemoveForceAlignButton()
        {
            foreach (SideGridItemRight obj in DocCompareSideListViewRight.Items)
            {
                obj.RemoveForceAlignButtonEnable = false;
            }
        }

        private void DisableSideScrollLeft()
        {
            //DocCompareSideScrollViewerLeft.IsEnabled = false;
        }

        private void DisableSideScrollRight()
        {
            //DocCompareSideScrollViewerRight.IsEnabled = false;
        }

        private void DisplayComparisonResult()
        {
            Dispatcher.Invoke(() =>
            {
                DocCompareNameLabel1.Text = Path.GetFileName(docs.documents[docs.documentsToCompare[0]].filePath);
            });

            List<CompareMainItem> mainItemList = new List<CompareMainItem>();

            for (int i = 0; i < docs.totalLen; i++)
            {
                CompareMainItem thisItem = new CompareMainItem()
                {
                    ImgGridName = "MainImgGrid" + i.ToString(),
                    ImgGridLeftName = "MainImgGridLeft" + i.ToString(),
                    ImgGridRightName = "MainImgGridRight" + i.ToString(),
                    ImgLeftName = "MainImgLeft" + i.ToString(),
                    ImgRightName = "MainImgRight" + i.ToString(),
                    ImgAniLeftName = "MainImgAniLeft" + i.ToString(),
                    ImgAniRightName = "MainImgAniRight" + i.ToString(),
                    ImgMaskRightName = "MainMaskImgRight" + i.ToString(),
                    AnimateDiffLeftButtonName = "AnimateDiffLeft" + i.ToString(),
                    AnimateDiffRightButtonName = "AnimateDiffRight" + i.ToString(),
                    AniDiffButtonEnable = false,
                    Margin = new Thickness(10),
                    PPTSpeakerNoteGridNameLeft = "PPTSpeakerNoteGridLeft" + i.ToString(),
                    PPTSpeakerNoteGridNameRight = "PPTSpeakerNoteGridRight" + i.ToString(),
                    ClosePPTSpeakerNotesButtonNameLeft = "ClosePPTSpeakerNotesButtonNameLeft" + i.ToString(),
                    ClosePPTSpeakerNotesButtonNameRight = "ClosePPTSpeakerNotesButtonNameRight" + i.ToString(),
                    ShowPPTSpeakerNotesButtonNameRight = "ShowPPTSpeakerNotesButtonNameRight" + i.ToString(),
                    ShowPPTSpeakerNotesButtonNameRightChanged = "ShowPPTSpeakerNotesButtonNameRightChanged" + i.ToString(),
                    ShowPPTSpeakerNotesButtonNameLeft = "ShowPPTSpeakerNotesButtonNameLeft" + i.ToString(),
                    PPTNoteGridLeftVisi = Visibility.Hidden,
                    PPTNoteGridRightVisi = Visibility.Hidden,
                    showPPTSpeakerNotesButtonRight = Visibility.Hidden,
                    showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden,
                };

                bool showSpeakerNotesLeft = false;
                bool showSpeakerNotesRight = false;
                bool didChange = false;

                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                {
                    thisItem.PathToImgLeft = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".png");

                    if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                    {
                        thisItem.PathToAniImgLeft = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");
                        thisItem.AniDiffButtonEnable = true;
                    }

                    if (docs.documents[docs.documentsToCompare[0]].fileType == Document.FileTypes.PPT)
                    {
                        if (docs.documents[docs.documentsToCompare[0]].pptIsHidden[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]] == true)
                        {
                            thisItem.ShowHiddenLeft = Visibility.Visible;
                            thisItem.BlurRadiusLeft = 5;
                        }
                        else
                        {
                            thisItem.ShowHiddenLeft = Visibility.Hidden;
                            thisItem.BlurRadiusLeft = 0;
                        }

                        if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                        {
                            thisItem.Document1 = "<?xml version=\"1.0\"?> \n<text>" + docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]] + "</text>";
                            showSpeakerNotesLeft = true;
                        }
                        else
                        {
                            thisItem.Document1 = "<?xml version=\"1.0\"?> \n" + "<text></text>";
                            if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                            {
                                if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes != null)
                                {
                                    if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                                    {
                                        didChange |= true;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        thisItem.ShowHiddenLeft = Visibility.Hidden;
                        thisItem.BlurRadiusLeft = 0;
                    }
                }
                else
                {
                    thisItem.ShowHiddenLeft = Visibility.Hidden;
                    thisItem.BlurRadiusLeft = 0;
                }

                if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                {
                    thisItem.PathToImgRight = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");

                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                    {
                        thisItem.PathToAniImgRight = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".png");

                        if (File.Exists(Path.Join(compareResultFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png")))
                        {
                            thisItem.PathToMaskImgRight = Path.Join(compareResultFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");
                            thisItem.AniDiffButtonEnable = true;

                            if (showMask == true)
                                thisItem.ShowMask = Visibility.Visible;
                            else
                                thisItem.ShowMask = Visibility.Hidden;
                        }
                    }

                    if (docs.documents[docs.documentsToCompare[1]].fileType == Document.FileTypes.PPT)
                    {
                        if (docs.documents[docs.documentsToCompare[1]].pptIsHidden[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]] == true)
                        {
                            thisItem.ShowHiddenRight = Visibility.Visible;
                            thisItem.BlurRadiusRight = 5;
                        }
                        else
                        {
                            thisItem.ShowHiddenRight = Visibility.Hidden;
                            thisItem.BlurRadiusRight = 0;
                        }


                        if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                        {
                            if (docs.pptSpeakerNotesDiff[i].Count == 0)
                            {
                                thisItem.Document2 = "<?xml version=\"1.0\"?> \n<text>" + docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]] + "</text>";
                                thisItem.showPPTSpeakerNotesButtonRight = Visibility.Visible;
                                thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;

                                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                                {
                                    if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes != null)
                                    {
                                        if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                                        {
                                            didChange |= true;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                string doc = "<?xml version =\"1.0\"?> \n<text>";


                                foreach (DocCompareDLL.Diff diff in docs.pptSpeakerNotesDiff[i])
                                {
                                    if (diff.operation == DocCompareDLL.Operation.INSERT)
                                    {
                                        doc += "<INSERT>" + diff.text + "</INSERT>";
                                        didChange |= true;
                                    }
                                    else if (diff.operation == DocCompareDLL.Operation.DELETE)
                                    {
                                        doc += "<DELETE>" + diff.text + "</DELETE>";
                                        didChange |= true;
                                    }
                                    else
                                    {
                                        doc += diff.text;
                                    }

                                }

                                doc += "</text>";

                                thisItem.Document2 = doc;
                            }

                            showSpeakerNotesRight = true;
                        }
                        else
                        {
                            thisItem.Document2 = "<?xml version=\"1.0\"?> \n" + "<text></text>";
                            thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                            thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                            //thisItem.ShowPPTNoteButtonBackground = Brushes.White;

                            if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                            {
                                if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes != null)
                                {
                                    if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                                    {
                                        didChange |= true;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        thisItem.ShowHiddenRight = Visibility.Hidden;
                        thisItem.BlurRadiusRight = 0;
                    }
                }
                else
                {
                    thisItem.ShowHiddenRight = Visibility.Hidden;
                    thisItem.BlurRadiusRight = 0;
                }

                if (showSpeakerNotesLeft == false)
                {
                    if (showSpeakerNotesRight == true)
                    {

                        thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Visible;
                        thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                    }
                    else
                    {
                        thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                    }
                    thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                }
                else
                {
                    /*if (showSpeakerNotesRight == true)
                    {
                        thisItem.showPPTSpeakerNotesButtonRight = Visibility.Visible;
                        thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                    }
                    
                    else
                    {
                    */
                    //thisItem.showPPTSpeakerNotesButtonRight = Visibility.Visible;

                    if (showSpeakerNotesRight == true)
                    {
                        if (didChange == true)
                        {
                            thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                            thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Visible;
                        }
                        else
                        {
                            thisItem.showPPTSpeakerNotesButtonRight = Visibility.Visible;
                            thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                        }
                        thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                    }
                    else
                    {
                        thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Visible;
                    }
                    //}
                }

                mainItemList.Add(thisItem);
            }

            Dispatcher.Invoke(() => { DocCompareMainListView.ItemsSource = mainItemList; });

            List<SideGridItemLeft> leftItemList = new List<SideGridItemLeft>();
            List<SideGridItemRight> rightItemList = new List<SideGridItemRight>();

            for (int i = 0; i < docs.totalLen; i++)
            {
                SideGridItemLeft leftItem = new SideGridItemLeft()
                {
                    GridName = "LeftSideGrid" + i.ToString(),
                    ImgGridName = "SideImageLeft" + i.ToString(),
                    ForceAlignButtonName = "SideButtonLeft" + i.ToString(),
                    ForceAlignInvalidButtonName = "SideButtonInvalidLeft" + i.ToString(),
                    Margin = new Thickness(10),
                    PageNumberLabel = (i + 1).ToString(),
                    BackgroundBrush = Color.FromArgb(0, 255, 255, 255),
                    ShowHidden = Visibility.Hidden,
                    ForceAlignButtonVisi = Visibility.Hidden,
                    ForceAlignButtonInvalidVisi = Visibility.Hidden,
                    RemoveForceAlignButtonName = "RemoveForceAlign" + i.ToString()
                };

                SideGridItemRight rightItem = new SideGridItemRight()
                {
                    GridName = "RightSideGrid" + i.ToString(),
                    ImgGridName = "SideImageRight" + i.ToString(),
                    ImgMaskName = "SideImageRightMask" + i.ToString(),
                    ForceAlignButtonName = "SideButtonRight" + i.ToString(),
                    ForceAlignInvalidButtonName = "SideButtonInvalidRight" + i.ToString(),
                    Margin = new Thickness(10),
                    RemoveForceAlignButtonEnable = false,
                    RemoveForceAlignButtonVisibility = Visibility.Hidden,
                    BackgroundBrush = Color.FromArgb(0, 255, 255, 255),
                    ShowHidden = Visibility.Hidden,
                    ForceAlignButtonVisi = Visibility.Hidden,
                    ForceAlignButtonInvalidVisi = Visibility.Hidden,
                    RemoveForceAlignButtonName = "RemoveForceAlign" + i.ToString(),
                };

                if (i == 0)
                {
                    leftItem.BackgroundBrush = Color.FromArgb(255, 119, 119, 119);
                    rightItem.BackgroundBrush = Color.FromArgb(255, 119, 119, 119);
                }

                // Remove force alignment button
                foreach (List<int> ind1 in docs.forceAlignmentIndices)
                {
                    if (ind1[0] == docs.documents[docs.documentsToCompare[0]].docCompareIndices[i])
                    {
                        /*
                        
                        rightItem.RemoveForceAlignButtonVisibility = Visibility.Visible;
                        rightItem.RemoveForceAlignButtonEnable = true;
                        */
                        leftItem.ForceAlignButtonVisi = Visibility.Visible;
                        rightItem.ForceAlignButtonVisi = Visibility.Visible;
                    }
                }

                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                {
                    leftItem.PathToImg = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".png");
                    rightItem.PathToImgDummy = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".png");

                    if (docs.documents[docs.documentsToCompare[0]].fileType == Document.FileTypes.PPT)
                    {
                        if (docs.documents[docs.documentsToCompare[0]].pptIsHidden[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]] == true)
                        {
                            leftItem.ShowHidden = Visibility.Visible;
                        }
                        else
                        {
                            leftItem.ShowHidden = Visibility.Hidden;
                        }
                    }
                    else
                    {
                        leftItem.ShowHidden = Visibility.Hidden;
                    }
                }

                if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1) // doc 2 has a valid page
                {
                    rightItem.PathToImg = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");
                    leftItem.PathToImgDummy = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");

                    if (File.Exists(Path.Join(compareResultFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png")))
                    {
                        rightItem.PathToMask = Path.Join(compareResultFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");
                        if (i != 0)
                        {
                            leftItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            rightItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            rightItem.DiffVisi = Visibility.Visible;
                            rightItem.NoDiffVisi = Visibility.Hidden;
                        }

                        if (showMask == true)
                            rightItem.ShowMask = Visibility.Visible;
                        else
                            rightItem.ShowMask = Visibility.Hidden;
                    }
                    else
                    {
                        rightItem.DiffVisi = Visibility.Hidden;
                        rightItem.NoDiffVisi = Visibility.Visible;
                    }

                    if (docs.documents[docs.documentsToCompare[1]].fileType == Document.FileTypes.PPT)
                    {
                        if (docs.documents[docs.documentsToCompare[1]].pptIsHidden[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]] == true)
                        {
                            rightItem.ShowHidden = Visibility.Visible;
                        }
                        else
                        {
                            rightItem.ShowHidden = Visibility.Hidden;
                        }
                    }
                    else
                    {
                        rightItem.ShowHidden = Visibility.Hidden;
                    }
                }

                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] == -1 || docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] == -1)
                {
                    rightItem.NoDiffVisi = Visibility.Hidden;
                    rightItem.DiffVisi = Visibility.Hidden;
                }

                leftItemList.Add(leftItem);
                rightItemList.Add(rightItem);
            }

            Dispatcher.Invoke(() =>
            {
                DocCompareSideListViewLeft.ItemsSource = leftItemList;
                DocCompareSideListViewRight.ItemsSource = rightItemList;
                //DocCompareSideScrollViewerLeft.Content = docCompareChildPanelLeft;
                //DocCompareSideScrollViewerRight.Content = docCompareChildPanelRight;

                docCompareGrid.Visibility = Visibility.Visible;
                ProgressBarDocCompare.Visibility = Visibility.Hidden;
                ProgressBarDocCompareAlign.Visibility = Visibility.Hidden;
                ProgressBarLoadingResults.Visibility = Visibility.Hidden;

                if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPARELINK)
                {
                    //ListViewItem container = DocCompareSideListViewLeft.ItemContainerGenerator.ContainerFromItem(DocCompareSideListViewLeft.Items[0]) as ListViewItem;
                    //(((VisualTreeHelper.GetChild(container, 0) as Grid).Children[1] as Grid).Children[3] as Button).Visibility = Visibility.Visible;
                    PopupLinkPageBubble.IsOpen = true;
                    //walkthroughStep++;
                }
            });
        }

        private void DisplayPreview(int viewerID, int docIndex)
        {
            if (docIndex != -1)
            {
                if (docs.documents.Count >= 1)
                {
                    List<SimpleImageItem> imageList = new List<SimpleImageItem>();
                    if (docs.documents[docIndex].filePath != null)
                    {
                        int pageCounter = 0;

                        DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                        FileInfo[] fi = di.GetFiles();

                        if (fi.Length != 0)
                        {
                            for (int i = 0; i < fi.Length; i++)
                            {
                                SimpleImageItem thisImage = new SimpleImageItem()
                                {
                                    PathToFile = Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".png"),
                                    EoDVisi = Visibility.Hidden
                                };

                                if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptIsHidden[i] == true)
                                {
                                    thisImage.BlurRadius = 5;
                                    thisImage.showHidden = Visibility.Visible;
                                    thisImage.HiddenPPTGridName = "HiddenPPTGrid" + i.ToString();
                                }
                                else
                                {
                                    thisImage.BlurRadius = 0;
                                    thisImage.showHidden = Visibility.Hidden;
                                }

                                if (pageCounter == 0)
                                {
                                    thisImage.Margin = new Thickness(10, 10, 10, 10);
                                }
                                else
                                {
                                    thisImage.Margin = new Thickness(10, 0, 10, 10);
                                }

                                if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptSpeakerNotes[i].Length != 0)
                                {
                                    thisImage.showPPTSpeakerNotesButton = Visibility.Visible;
                                    thisImage.ShowPPTSpeakerNotesButtonName = "ShowPPTNoteButton" + i.ToString();
                                    thisImage.PPTSpeakerNoteGridName = "PPTSpeakerNoteGrid" + i.ToString();
                                    thisImage.ClosePPTSpeakerNotesButtonName = "ClosePPTSpeakerNoteButton" + i.ToString();
                                    thisImage.PPTSpeakerNotes = docs.documents[docIndex].pptSpeakerNotes[i];
                                }
                                else
                                {
                                    thisImage.showPPTSpeakerNotesButton = Visibility.Hidden;
                                }

                                imageList.Add(thisImage);
                                pageCounter++;
                            }
                        }

                        // add End of Document
                        SimpleImageItem item = new SimpleImageItem
                        {
                            EoDVisi = Visibility.Visible,
                            showHidden = Visibility.Hidden,
                            showPPTSpeakerNotesButton = Visibility.Hidden,
                        };
                        imageList.Add(item);
                    }

                    if (imageList.Count != 0)
                    {
                        switch (viewerID)
                        {
                            case 1:
                                DocCompareListView1.ItemsSource = imageList;
                                DocCompareListView1.Items.Refresh();
                                DocCompareListView1.ScrollIntoView(DocCompareListView1.Items[0]);
                                Doc1Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc1.Visibility = Visibility.Hidden;
                                //Doc1PageNumberLabel.Content = "1 / " + DocCompareListView1.Items.Count.ToString();
                                break;
                            case 2:
                                DocCompareListView2.ItemsSource = imageList;
                                DocCompareListView2.Items.Refresh();
                                DocCompareListView2.ScrollIntoView(DocCompareListView2.Items[0]);
                                Doc2Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc2.Visibility = Visibility.Hidden;
                                Doc2NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[1]].filePath);
                                //Doc2PageNumberLabel.Content = "1 / " + DocCompareListView2.Items.Count.ToString();
                                break;
                            case 3:
                                DocCompareListView3.ItemsSource = imageList;
                                DocCompareListView3.Items.Refresh();
                                DocCompareListView3.ScrollIntoView(DocCompareListView3.Items[0]);
                                Doc3Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc3.Visibility = Visibility.Hidden;
                                Doc3NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[2]].filePath);
                                //Doc3PageNumberLabel.Content = "1 / " + DocCompareListView3.Items.Count.ToString();
                                break;
                            case 4:
                                DocCompareListView4.ItemsSource = imageList;
                                DocCompareListView4.Items.Refresh();
                                DocCompareListView4.ScrollIntoView(DocCompareListView4.Items[0]);
                                Doc4Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc4.Visibility = Visibility.Hidden;
                                Doc4NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[3]].filePath);
                                //Doc4PageNumberLabel.Content = "1 / " + DocCompareListView4.Items.Count.ToString();
                                break;
                            case 5:
                                DocCompareListView5.ItemsSource = imageList;
                                DocCompareListView5.Items.Refresh();
                                DocCompareListView5.ScrollIntoView(DocCompareListView5.Items[0]);
                                Doc5Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc5.Visibility = Visibility.Hidden;
                                Doc5NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[4]].filePath);
                                //Doc5PageNumberLabel.Content = "1 / " + DocCompareListView5.Items.Count.ToString();
                                break;
                        }
                    }
                }
            }
        }

        private void DisplayImageLeft(int docIndex)
        {
            if (docIndex != -1)
            {
                if (docs.documents.Count >= 1)
                {
                    if (docs.documents[docIndex].filePath != null)
                    {
                        int pageCounter = 0;
                        List<SimpleImageItem> imageList = new List<SimpleImageItem>();

                        DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                        FileInfo[] fi = di.GetFiles();

                        if (fi.Length != 0)
                        {
                            for (int i = 0; i < fi.Length; i++)
                            {
                                SimpleImageItem thisImage = new SimpleImageItem()
                                {
                                    PathToFile = Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".png"),
                                    EoDVisi = Visibility.Hidden
                                };

                                if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptIsHidden[i] == true)
                                {
                                    thisImage.BlurRadius = 5;
                                    thisImage.showHidden = Visibility.Visible;
                                    thisImage.HiddenPPTGridName = "HiddenPPTGrid" + i.ToString();
                                }
                                else
                                {
                                    thisImage.BlurRadius = 0;
                                    thisImage.showHidden = Visibility.Hidden;
                                }

                                if (pageCounter == 0)
                                {
                                    thisImage.Margin = new Thickness(10, 10, 10, 10);
                                }
                                else
                                {
                                    thisImage.Margin = new Thickness(10, 0, 10, 10);
                                }

                                if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptSpeakerNotes[i].Length != 0)
                                {
                                    thisImage.showPPTSpeakerNotesButton = Visibility.Visible;
                                    thisImage.ShowPPTSpeakerNotesButtonName = "ShowPPTNoteButton" + i.ToString();
                                    thisImage.PPTSpeakerNoteGridName = "PPTSpeakerNoteGrid" + i.ToString();
                                    thisImage.ClosePPTSpeakerNotesButtonName = "ClosePPTSpeakerNoteButton" + i.ToString();
                                    thisImage.PPTSpeakerNotes = docs.documents[docIndex].pptSpeakerNotes[i];
                                }
                                else
                                {
                                    thisImage.showPPTSpeakerNotesButton = Visibility.Hidden;
                                }

                                imageList.Add(thisImage);
                                pageCounter++;
                            }
                        }

                        // add End of Document
                        SimpleImageItem item = new SimpleImageItem
                        {
                            EoDVisi = Visibility.Visible,
                            showHidden = Visibility.Hidden,
                            showPPTSpeakerNotesButton = Visibility.Hidden,
                        };
                        imageList.Add(item);

                        DocCompareListView1.ItemsSource = imageList;
                        DocCompareListView1.Items.Refresh();
                        DocCompareListView1.ScrollIntoView(DocCompareListView1.Items[0]);
                        Doc1Grid.Visibility = Visibility.Visible;
                        ProgressBarDoc1.Visibility = Visibility.Hidden;
                        //Doc1PageNumberLabel.Content = "1 / " + DocCompareListView1.Items.Count.ToString();
                    }
                }
            }
        }

        private void DisplayImageMiddle(int docIndex)
        {
            if (docIndex != -1)
            {
                Brush brush = FindResource("DocumentBackGroundBrush") as Brush;
                Dispatcher.Invoke(() =>
                {
                    if (docs.documents.Count >= 2)
                    {
                        if (docs.documents[docIndex].filePath != null)
                        {
                            int pageCounter = 0;
                            List<SimpleImageItem> imageList = new List<SimpleImageItem>();

                            DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                            FileInfo[] fi = di.GetFiles();

                            if (fi.Length != 0)
                            {
                                for (int i = 0; i < fi.Length; i++)
                                {
                                    SimpleImageItem thisImage = new SimpleImageItem()
                                    {
                                        PathToFile = Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".png"),
                                        EoDVisi = Visibility.Hidden
                                    };

                                    if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptIsHidden[i] == true)
                                    {
                                        thisImage.BlurRadius = 5;
                                        thisImage.showHidden = Visibility.Visible;
                                        thisImage.HiddenPPTGridName = "HiddenPPTGrid" + i.ToString();
                                    }
                                    else
                                    {
                                        thisImage.BlurRadius = 0;
                                        thisImage.showHidden = Visibility.Hidden;
                                    }

                                    if (pageCounter == 0)
                                    {
                                        thisImage.Margin = new Thickness(10, 10, 10, 10);
                                    }
                                    else
                                    {
                                        thisImage.Margin = new Thickness(10, 0, 10, 10);
                                    }

                                    if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptSpeakerNotes[i].Length != 0)
                                    {
                                        thisImage.showPPTSpeakerNotesButton = Visibility.Visible;
                                        thisImage.ShowPPTSpeakerNotesButtonName = "ShowPPTNoteButton" + i.ToString();
                                        thisImage.PPTSpeakerNoteGridName = "PPTSpeakerNoteGrid" + i.ToString();
                                        thisImage.ClosePPTSpeakerNotesButtonName = "ClosePPTSpeakerNoteButton" + i.ToString();
                                        thisImage.PPTSpeakerNotes = docs.documents[docIndex].pptSpeakerNotes[i];
                                    }
                                    else
                                    {
                                        thisImage.showPPTSpeakerNotesButton = Visibility.Hidden;
                                    }

                                    imageList.Add(thisImage);
                                    pageCounter++;
                                }
                            }

                            // add End of Document
                            SimpleImageItem item = new SimpleImageItem
                            {
                                EoDVisi = Visibility.Visible,
                                showHidden = Visibility.Hidden,
                                showPPTSpeakerNotesButton = Visibility.Hidden,
                            };
                            imageList.Add(item);

                            DocCompareListView2.ItemsSource = imageList;
                            DocCompareListView2.Items.Refresh();
                            DocCompareListView2.ScrollIntoView(DocCompareListView2.Items[0]);
                            Doc2Grid.Visibility = Visibility.Visible;
                            ProgressBarDoc2.Visibility = Visibility.Hidden;
                            //Doc2PageNumberLabel.Content = "1 / " + DocCompareListView2.Items.Count.ToString();
                        }
                    }
                });
            }
        }

        private void DisplayImageRight(int docIndex)
        {
            if (docIndex != -1)
            {
                Brush brush = FindResource("DocumentBackGroundBrush") as Brush;
                Dispatcher.Invoke(() =>
                {
                    if (docs.documents.Count >= 3)
                    {
                        if (docs.documents[docIndex].filePath != null)
                        {
                            int pageCounter = 0;
                            List<SimpleImageItem> imageList = new List<SimpleImageItem>();

                            DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                            FileInfo[] fi = di.GetFiles();

                            if (fi.Length != 0)
                            {
                                for (int i = 0; i < fi.Length; i++)
                                {
                                    SimpleImageItem thisImage = new SimpleImageItem()
                                    {
                                        PathToFile = Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".png")
                                    };

                                    if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptIsHidden[i] == true)
                                    {
                                        thisImage.BlurRadius = 5;
                                        thisImage.showHidden = Visibility.Visible;
                                        thisImage.HiddenPPTGridName = "HiddenPPTGrid" + i.ToString();
                                    }
                                    else
                                    {
                                        thisImage.BlurRadius = 0;
                                        thisImage.showHidden = Visibility.Hidden;
                                    }

                                    if (pageCounter == 0)
                                    {
                                        thisImage.Margin = new Thickness(10, 10, 10, 10);
                                    }
                                    else
                                    {
                                        thisImage.Margin = new Thickness(10, 0, 10, 10);
                                    }

                                    if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptSpeakerNotes[i].Length != 0)
                                    {
                                        thisImage.showPPTSpeakerNotesButton = Visibility.Visible;
                                        thisImage.ShowPPTSpeakerNotesButtonName = "ShowPPTNoteButton" + i.ToString();
                                        thisImage.PPTSpeakerNoteGridName = "PPTSpeakerNoteGrid" + i.ToString();
                                        thisImage.ClosePPTSpeakerNotesButtonName = "ClosePPTSpeakerNoteButton" + i.ToString();
                                        thisImage.PPTSpeakerNotes = docs.documents[docIndex].pptSpeakerNotes[i];
                                    }
                                    else
                                    {
                                        thisImage.showPPTSpeakerNotesButton = Visibility.Hidden;
                                    }

                                    imageList.Add(thisImage);
                                    pageCounter++;
                                }
                            }

                            SimpleImageItem item = new SimpleImageItem
                            {
                                EoDVisi = Visibility.Visible,
                                showHidden = Visibility.Hidden,
                                showPPTSpeakerNotesButton = Visibility.Hidden,
                            };
                            imageList.Add(item);

                            DocCompareListView3.ItemsSource = imageList;
                            DocCompareListView3.ScrollIntoView(DocCompareListView2.Items[0]);
                            Doc3Grid.Visibility = Visibility.Visible;
                            ProgressBarDoc3.Visibility = Visibility.Hidden;
                        }
                    }
                });
            }
        }

        private void DisplayLicense()
        {
            switch (lic.GetLicenseTypes())
            {
                case LicenseManagement.LicenseTypes.ANNUAL_SUBSCRIPTION:
                    LicenseTypeLabel.Content = "Annual subscription";
                    LicenseExpiryTypeLabel.Content = "Renewal on";
                    LicenseExpiryLabel.Content = lic.GetExpiryDateString();
                    UserEmailTextBox.Text = lic.GetEmail();
                    LicenseKeyTextBox.Text = lic.GetKey();
                    UserEmailTextBox.IsEnabled = false;
                    LicenseKeyTextBox.IsEnabled = false;
                    ActivateLicenseButton.Visibility = Visibility.Hidden;
                    ActivateLicenseButton.IsEnabled = false;
                    ChangeLicenseButton.Visibility = Visibility.Visible;
                    ChangeLicenseButton.IsEnabled = true;
                    ExtendTrialGrid1.Visibility = Visibility.Hidden;
                    ExtendTrialGrid2.Visibility = Visibility.Hidden;
                    break;

                case LicenseManagement.LicenseTypes.TRIAL:
                    LicenseTypeLabel.Content = "Trial license";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    TimeSpan timeBuffer = lic.GetExpiryDate().Subtract(DateTime.Today);
                    if (timeBuffer.TotalDays >= 0)
                        LicenseExpiryLabel.Content = timeBuffer.TotalDays.ToString() + " days";
                    else
                        LicenseExpiryTypeLabel.Content = "Expired";

                    if (settings.showExtendTrial == true && settings.trialExtended == false)
                    {
                        ExtendTrialGrid1.Visibility = Visibility.Visible;
                        ExtendTrialGrid2.Visibility = Visibility.Visible;
                    }

                    break;

                case LicenseManagement.LicenseTypes.DEVELOPMENT:
                    LicenseTypeLabel.Content = "Developer license";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    LicenseExpiryLabel.Content = "- days";

                    ExtendTrialGrid1.Visibility = Visibility.Hidden;
                    ExtendTrialGrid2.Visibility = Visibility.Hidden;
                    break;

                default:
                    LicenseTypeLabel.Content = "No license found";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    LicenseExpiryLabel.Content = "- days";

                    ExtendTrialGrid1.Visibility = Visibility.Hidden;
                    ExtendTrialGrid2.Visibility = Visibility.Hidden;
                    break;
            }

            switch (lic.GetLicenseStatus())
            {
                case LicenseManagement.LicenseStatus.ACTIVE:
                    LicenseStatusTypeLabel.Content = "License status";
                    LicenseStatusLabel.Content = "Active";
                    BrowseFileButton1.IsEnabled = true;
                    DocCompareFirstDocZone.AllowDrop = true;
                    DocCompareDragDropZone1.AllowDrop = true;
                    DocCompareColorZone1.AllowDrop = true;
                    break;

                case LicenseManagement.LicenseStatus.INACTIVE:
                    LicenseStatusTypeLabel.Content = "License status";
                    LicenseStatusLabel.Content = "Inactive";
                    BrowseFileButton1.IsEnabled = false;
                    DocCompareFirstDocZone.AllowDrop = false;
                    DocCompareDragDropZone1.AllowDrop = false;
                    DocCompareColorZone1.AllowDrop = false;
                    break;
            }
        }

        private void DisplayRefDoc(int docIndex)
        {
            Brush brush = FindResource("DocumentBackGroundBrush") as Brush;
            Dispatcher.Invoke(() =>
            {
                if (docs.documents[docIndex].filePath != null)
                {
                    int pageCounter = 0;
                    refDocPanel = new StackPanel
                    {
                        Background = brush,
                        HorizontalAlignment = HorizontalAlignment.Stretch
                    };

                    DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                    FileInfo[] fi = di.GetFiles();

                    for (int i = 0; i < fi.Length; i++)
                    {
                        Image thisImage = new Image();
                        var stream = File.OpenRead(Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".png"));
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = stream;
                        bitmap.EndInit();
                        bitmap.Freeze();
                        stream.Close();
                        thisImage.Source = bitmap;
                        if (pageCounter == 0)
                        {
                            thisImage.Margin = new Thickness(10, 10, 10, 10);
                        }
                        else
                        {
                            thisImage.Margin = new Thickness(10, 0, 10, 10);
                        }

                        thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        refDocPanel.Children.Add(thisImage);
                        pageCounter++;
                    }

                    RefDocScrollViewer.Content = refDocPanel;
                    RefDocScrollViewer.ScrollToVerticalOffset(0);
                }
            });
        }

        private void Doc1NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc1NameLabelComboBox.SelectedItem.ToString();
                int desiredInd = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);

                int existingInd = -1;
                for (int i = 0; i < docs.documents.Count; i++)
                {
                    if (docs.documentsToShow[i] == desiredInd)
                    {
                        existingInd = i;
                    }
                }

                docs.documentsToShow[existingInd] = docs.documentsToShow[0];
                docs.documentsToShow[0] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                //DisplayImageLeft(docs.documentsToShow[0]);
                DisplayPreview(1, docs.documentsToShow[0]);
                DisplayPreview(existingInd + 1, docs.documentsToShow[existingInd]);
                UpdateDocSelectionComboBox();
                UpdateFileStat(0);
            }
            catch
            {
                Doc1NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
                UpdateFileStat(0);
            }
        }
        /*
        private void Doc2NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc2NameLabelComboBox.SelectedItem.ToString();
                docs.documentsToShow[1] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayPreview(2, docs.documentsToShow[1]);
                UpdateDocSelectionComboBox();
                UpdateFileStat(1);
            }
            catch
            {
                Doc2NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
                UpdateFileStat(1);
            }
        }

        private void Doc3NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc3NameLabelComboBox.SelectedItem.ToString();
                docs.documentsToShow[2] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayPreview(3, docs.documentsToShow[2]);
                UpdateDocSelectionComboBox();
                UpdateFileStat(2);
            }
            catch
            {
                Doc3NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
                UpdateFileStat(2);
            }
        }
        */
        private void DocCompareDragDropZone1_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void DocCompareDragDropZone1_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON1)
                {
                    PopupBrowseFileButtonBubble.IsOpen = false;
                }

                var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                string ext;

                foreach (string file in data)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    SelectReferenceWindow selectReferenceWindow = new SelectReferenceWindow();
                    List<string> fileList = new List<string>();
                    foreach (Document doc in docs.documents)
                    {
                        fileList.Add(Path.GetFileName(doc.filePath));
                    }

                    selectReferenceWindow.Setup(fileList);

                    if (docs.documents.Count >= 2)
                    {
                        if (selectReferenceWindow.ShowDialog() == true)
                        {
                            docs.documentsToShow[0] = selectReferenceWindow.selectedIndex;
                        }
                    }

                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();

                    if (walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON1)
                    {
                        walkthroughStep = WalkthroughSteps.BROWSEFILEBUTTON2;
                    }
                }
                else
                {
                    if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON1)
                    {
                        PopupBrowseFileButtonBubble.IsOpen = true;
                    }
                }
            }
        }

        private void DocCompareDragDropZone2_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void DocCompareDragDropZone2_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                string ext;

                if (data.Length > settings.maxDocCount)
                    ShowMaxDocCountWarningBox();

                foreach (string file in data)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    LoadFilesCommonPart();

                    //docs.documentsToShow[1] = docs.documents.Count - 1;

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }

        private void DocCompareDragDropZone3_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void DocCompareDragDropZone3_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                string ext;

                foreach (string file in data)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    LoadFilesCommonPart();

                    //docs.documentsToShow[2] = docs.documents.Count - 1;

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }

        private void DocCompareMainScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                double accuHeight = 0;

                Border border = (Border)VisualTreeHelper.GetChild(DocCompareMainListView, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                for (int i = 0; i < DocCompareMainListView.Items.Count; i++)
                {
                    ListViewItem container = DocCompareMainListView.ItemContainerGenerator.ContainerFromItem(DocCompareMainListView.Items[i]) as ListViewItem;
                    accuHeight += container.ActualHeight;

                    if (docs.documents.Count >= 2)
                    {
                        if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3)
                        {
                            if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                            {
                                DocComparePageNumberLabel1.Content = (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] + 1).ToString() + " / " +
                                    (docs.documents[docs.documentsToCompare[0]].docCompareIndices.Max() + 1).ToString();
                            }

                            if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                            {
                                DocComparePageNumberLabel2.Content = (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] + 1).ToString() + " / " +
                                    (docs.documents[docs.documentsToCompare[1]].docCompareIndices.Max() + 1).ToString();
                            }

                            docCompareSideGridShown = i;
                            HighlightSideGrid();
                            break;
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void DocCompareNameLabel2ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = DocCompareNameLabel2ComboBox.SelectedItem.ToString();
                if (docs.documentsToCompare[1] != docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName))
                {
                    docs.documentsToCompare[1] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                    docs.forceAlignmentIndices = new List<List<int>>();
                }

                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ProgressBarDocCompare.Visibility = Visibility.Visible;
                });

                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();

                UpdateFileStat(5);
            }
            catch
            {
            }
        }

        private void DocCompareScrollViewer1_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                double accuHeight = 0;

                Border border = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                for (int i = 0; i < DocCompareListView1.Items.Count - 1; i++)
                {
                    ListViewItem container = DocCompareListView1.ItemContainerGenerator.ContainerFromItem(DocCompareListView1.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3 && Doc1Grid.Visibility == Visibility.Visible)
                    {
                        //Doc1PageNumberLabel.Content = (i + 1).ToString() + " / " + (DocCompareListView1.Items.Count - 1).ToString();
                        break;
                    }
                    else
                    {
                        //Doc1PageNumberLabel.Content = "";
                    }
                }

                if (linkscroll == true)
                {
                    // try to scroll others
                    Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                    ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }
                }
            }
            catch
            {
            }
        }

        private void DocCompareScrollViewer2_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                double accuHeight = 0;

                Border border = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                for (int i = 0; i < DocCompareListView2.Items.Count - 1; i++)
                {
                    ListViewItem container = DocCompareListView2.ItemContainerGenerator.ContainerFromItem(DocCompareListView2.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3 && Doc2Grid.Visibility == Visibility.Visible)
                    {
                        //Doc2PageNumberLabel.Content = (i + 1).ToString() + " / " + (DocCompareListView2.Items.Count -1 ).ToString();
                        break;
                    }
                    else
                    {
                        //Doc2PageNumberLabel.Content = "";
                    }
                }

                if (linkscroll == true)
                {
                    // try to scroll others
                    Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                    ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }
                }
            }
            catch
            {
            }
        }

        private void DocCompareScrollViewer3_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                double accuHeight = 0;

                Border border = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                for (int i = 0; i < DocCompareListView3.Items.Count - 1; i++)
                {
                    ListViewItem container = DocCompareListView3.ItemContainerGenerator.ContainerFromItem(DocCompareListView3.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3 && Doc2Grid.Visibility == Visibility.Visible)
                    {
                        //Doc3PageNumberLabel.Content = (i + 1).ToString() + " / " + (DocCompareListView3.Items.Count-1).ToString();
                        break;
                    }
                    else
                    {
                        //Doc3PageNumberLabel.Content = "";
                    }
                }

                if (linkscroll == true)
                {
                    // try to scroll others
                    Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                    ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }
                }
            }
            catch
            {
            }
        }

        private void DocCompareSideScrollViewerLeft_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                Border border = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewLeft, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewRight, 0);
                ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                if (inForceAlignMode == false)
                    scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                else
                {
                    if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                    {
                        scrollViewer.ScrollToVerticalOffset(scrollPosLeft);
                    }
                }
            }
            catch
            {
            }
        }

        private void DocCompareSideScrollViewerRight_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                Border border = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewLeft, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewRight, 0);
                ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                if (inForceAlignMode == false)
                    scrollViewer.ScrollToVerticalOffset(scrollViewer2.VerticalOffset);
                else
                {
                    if (sideGridSelectedLeftOrRight == GridSelection.RIGHT)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollPosRight);
                    }
                }
            }
            catch
            {
            }
        }

        private void EnableRemoveForceAlignButton()
        {
            foreach (SideGridItemRight obj in DocCompareSideListViewRight.Items)
            {
                obj.RemoveForceAlignButtonEnable = true;
            }
        }

        private void EnableSideScrollLeft()
        {
            //DocCompareSideScrollViewerLeft.IsEnabled = true;
        }

        private void EnableSideScrollRight()
        {
            //DocCompareSideScrollViewerRight.IsEnabled = true;
        }

        private int FindNextDocToShow()
        {
            for (int i = 0; i < docs.documents.Count; i++)
            {
                if (docs.documents[i].processed == true)
                {
                    if (i != docs.documentsToShow[0] && i != docs.documentsToShow[1])
                    {
                        if (docs.documentsToShow.Count == 3)
                        {
                            if (i != docs.documentsToShow[2])
                            {
                                return i;
                            }
                        }

                        return i;
                    }
                }
            }

            return -1;
        }

        private void HandleMainDocCompareAnimateMouseDown(object sender, MouseEventArgs args)
        {
            if (sender is Button)
            {
                string[] splittedName;
                if (mainGridSelectedLeftOrRight == GridSelection.LEFT)
                {
                    splittedName = (sender as Button).Tag.ToString().Split("Left");
                }
                else
                {
                    splittedName = (sender as Button).Tag.ToString().Split("Right");
                }

                gridToAnimate = (sender as Button).Parent as Grid;
                animateDiffRunning = true;

                foreach (CompareMainItem item in DocCompareMainListView.Items)
                {
                    if (item.AnimateDiffRightButtonName == (sender as Button).Tag.ToString())
                    {
                        item.ShowMask = Visibility.Hidden;
                        item.BlurRadiusRight = 0;
                        item.ShowHiddenRight = Visibility.Hidden;
                    }
                }

                threadAnimateDiff = new Thread(new ThreadStart(AnimateDiffThread));
                threadAnimateDiff.Start();

                if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPAREANIMATE)
                {
                    PopupAnimateBubble.IsOpen = false;
                }
            }
        }

        private void HandleMainDocCompareAnimateMouseRelease(object sender, MouseEventArgs args)
        {
            animateDiffRunning = false;
            if (sender is Button)
            {
                Grid parentGrid = (sender as Button).Parent as Grid;
                foreach (object child in parentGrid.Children)
                {
                    if (child is Border)
                    {
                        Image thisImg = (child as Border).Child as Image;
                        if (thisImg.Tag.ToString().Contains("Ani"))
                            thisImg.Visibility = Visibility.Hidden;
                        else if (thisImg.Tag.ToString().Contains("Mask"))
                        {
                        }
                        else
                            thisImg.Visibility = Visibility.Visible;
                    }
                }
            }

            for (int i = 0; i < DocCompareMainListView.Items.Count; i++)
            {
                CompareMainItem item = (CompareMainItem)DocCompareMainListView.Items[i];

                if (showMask == true)
                {
                    item.ShowMask = Visibility.Visible;
                }
                else
                {
                    item.ShowMask = Visibility.Hidden;
                }

                if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                {
                    if (docs.documents[docs.documentsToCompare[1]].fileType == Document.FileTypes.PPT)
                    {
                        if (docs.documents[docs.documentsToCompare[1]].pptIsHidden[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]] == true)
                        {
                            item.ShowHiddenRight = Visibility.Visible;
                            item.BlurRadiusRight = 5;
                        }
                        else
                        {
                            item.ShowHiddenRight = Visibility.Hidden;
                            item.BlurRadiusRight = 0;
                        }
                    }
                    else
                    {
                        item.ShowHiddenRight = Visibility.Hidden;
                        item.BlurRadiusRight = 0;
                    }
                }

            }

            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPAREANIMATE)
            {
                CustomMessageBox msgBox = new CustomMessageBox();
                msgBox.Setup("Walkthrough completed", "Thanks for completing the walkthrough guide.                 Enjoy 2|Compare! You can restart this walkthrough from     the settings page.", "Okay");
                msgBox.ShowDialog();

                settings.shownWalkthrough = true;
                SaveSettings();
                walkthroughStep = WalkthroughSteps.END;
            }
        }

        private void HandleMainDocCompareGridMouseEnter(object sender, MouseEventArgs args)
        {
            if (sender is Grid)
            {
                Grid parentGrid = sender as Grid;

                if (parentGrid.Tag.ToString().Contains("Left"))
                    mainGridSelectedLeftOrRight = GridSelection.LEFT;
                else
                    mainGridSelectedLeftOrRight = GridSelection.RIGHT;

                foreach (object child in parentGrid.Children)
                {
                    if (child is Button)
                    {
                        Button thisButton = child as Button;

                        if (mainGridSelectedLeftOrRight == GridSelection.LEFT)
                        {
                            if (thisButton.Tag != null)
                            {
                                if (thisButton.Tag.ToString().Contains("AnimateDiffLeft"))
                                    thisButton.Visibility = Visibility.Hidden;
                            }
                        }
                        else
                        {
                            if (thisButton.Tag != null)
                            {
                                if (thisButton.Tag.ToString().Contains("AnimateDiffRight"))
                                    thisButton.Visibility = Visibility.Visible;
                            }
                        }
                    }


                }

                Grid item = sender as Grid;
                Border childBorder = item.Children[0] as Border;
                Image img = childBorder.Child as Image;
                hiddenPPTEffect = img.Effect;
                img.Effect = null;

                if (parentGrid.Tag.ToString().Contains("Left"))
                {
                    hiddenPPTVisi = (item.Children[4] as Label).Visibility;
                    System.Windows.Shapes.Path path = item.Children[3] as System.Windows.Shapes.Path;
                    path.Visibility = Visibility.Hidden;
                    Label label = item.Children[4] as Label;
                    label.Visibility = Visibility.Hidden;
                    Grid grid = item.Children[2] as Grid;
                    grid.Visibility = Visibility.Hidden;
                }
                else
                {
                    hiddenPPTVisi = (item.Children[5] as Label).Visibility;
                    System.Windows.Shapes.Path path = item.Children[4] as System.Windows.Shapes.Path;
                    path.Visibility = Visibility.Hidden;
                    Label label = item.Children[5] as Label;
                    label.Visibility = Visibility.Hidden;
                    Grid grid = item.Children[3] as Grid;
                    grid.Visibility = Visibility.Hidden;
                }
            }
        }

        private void HandleMainDocCompareGridMouseLeave(object sender, MouseEventArgs args)
        {
            if (sender is Grid)
            {
                Grid parentGrid = sender as Grid;

                if (parentGrid.Tag.ToString().Contains("Left"))
                    mainGridSelectedLeftOrRight = GridSelection.LEFT;
                else
                    mainGridSelectedLeftOrRight = GridSelection.RIGHT;

                foreach (object child in parentGrid.Children)
                {
                    if (child is Button)
                    {
                        Button thisButton = child as Button;

                        if (mainGridSelectedLeftOrRight == GridSelection.LEFT)
                        {
                            if (thisButton.Tag != null)
                            {
                                if (thisButton.Tag.ToString().Contains("AnimateDiffLeft"))
                                    thisButton.Visibility = Visibility.Hidden;
                            }
                        }
                        else
                        {
                            if (thisButton.Tag != null)
                            {
                                if (thisButton.Tag.ToString().Contains("AnimateDiffRight"))
                                    thisButton.Visibility = Visibility.Hidden;
                            }
                        }
                    }
                }

                Grid item = sender as Grid;
                Border childBorder = item.Children[0] as Border;
                Image img = childBorder.Child as Image;
                img.Effect = hiddenPPTEffect;

                if (parentGrid.Tag.ToString().Contains("Left"))
                {
                    System.Windows.Shapes.Path path = item.Children[3] as System.Windows.Shapes.Path;
                    path.Visibility = hiddenPPTVisi;
                    Label label = item.Children[4] as Label;
                    label.Visibility = hiddenPPTVisi;
                    Grid grid = item.Children[2] as Grid;
                    grid.Visibility = hiddenPPTVisi;
                }
                else
                {
                    System.Windows.Shapes.Path path = item.Children[4] as System.Windows.Shapes.Path;
                    path.Visibility = hiddenPPTVisi;
                    Label label = item.Children[5] as Label;
                    label.Visibility = hiddenPPTVisi;
                    Grid grid = item.Children[3] as Grid;
                    grid.Visibility = hiddenPPTVisi;
                }
            }
        }

        private void HandleMouseClickOnSideScrollView(object sender, MouseButtonEventArgs e)
        {
            if (inForceAlignMode == false)
            {
                Border border = (Border)VisualTreeHelper.GetChild(DocCompareMainListView, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                Grid grid = sender as Grid;
                string[] splittedName = grid.Tag.ToString().Split("SideGrid");

                double accuHeight = 0;
                for (int i = 0; i < int.Parse(splittedName[1]); i++)
                {
                    ListViewItem container = DocCompareMainListView.ItemContainerGenerator.ContainerFromItem(DocCompareMainListView.Items[i]) as ListViewItem;
                    accuHeight += container.ActualHeight;
                }

                scrollViewer.ScrollToVerticalOffset(accuHeight);

                docCompareSideGridShown = int.Parse(splittedName[1]);
                HighlightSideGrid();
            }
        }

        private void HideDragDropZone2()
        {
            DocCompareSecondDocZone.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[1].Width = new GridLength(0, GridUnitType.Star);
            //Doc2PageNumberLabel.Visibility = Visibility.Collapsed;
            Doc2StatsGrid.Visibility = Visibility.Collapsed;
            ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
            DocPreviewStatGrid.ColumnDefinitions[1].Width = new GridLength(0, GridUnitType.Star);
        }

        private void HideDragDropZone3()
        {
            DocCompareThirdDocZone.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[2].Width = new GridLength(0, GridUnitType.Star);
            //Doc3PageNumberLabel.Visibility = Visibility.Collapsed;
            Doc3StatsGrid.Visibility = Visibility.Collapsed;
            ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
            DocPreviewStatGrid.ColumnDefinitions[2].Width = new GridLength(0, GridUnitType.Star);
        }

        private void HideDragDropZone4()
        {
            DocCompareFourthDocZone.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[3].Width = new GridLength(0, GridUnitType.Star);
            //Doc4PageNumberLabel.Visibility = Visibility.Collapsed;
            Doc4StatsGrid.Visibility = Visibility.Collapsed;
            ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
            DocPreviewStatGrid.ColumnDefinitions[3].Width = new GridLength(0, GridUnitType.Star);
        }

        private void HideDragDropZone5()
        {
            DocCompareFifthDocZone.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[4].Width = new GridLength(0, GridUnitType.Star);
            //Doc5PageNumberLabel.Visibility = Visibility.Collapsed;
            Doc5StatsGrid.Visibility = Visibility.Collapsed;
            ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
            DocPreviewStatGrid.ColumnDefinitions[4].Width = new GridLength(0, GridUnitType.Star);
        }

        private void HideMaskButton_Click(object sender, RoutedEventArgs e)
        {
            showMask = false;

            foreach (CompareMainItem item in DocCompareMainListView.Items)
            {
                item.ShowMask = Visibility.Hidden;
            }

            foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
            {
                item.ShowMask = Visibility.Hidden;
            }

            ShowMaskButton.Visibility = Visibility.Visible;
            HideMaskButton.Visibility = Visibility.Hidden;
            HighlightingDisableTip.Visibility = Visibility.Visible;

            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPAREHIGHLIGHT)
            {
                PopupHighlightOffBubble.IsOpen = false;
                PopupOpenOriBubble.IsOpen = true;
                walkthroughStep = WalkthroughSteps.COMPAREOPENEXTERN;
            }
        }

        private void HighlightSideGrid()
        {
            try
            {
                double accuHeight = 0;
                Border border = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewRight, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;
                double windowsHeight = scrollViewer.ActualHeight;

                for (int i = 0; i < DocCompareSideListViewLeft.Items.Count; i++)
                {
                    if (i == docCompareSideGridShown)
                    {
                        (DocCompareSideListViewLeft.Items[i] as SideGridItemLeft).BackgroundBrush = Color.FromArgb(255, 119, 119, 119);
                        (DocCompareSideListViewRight.Items[i] as SideGridItemRight).BackgroundBrush = Color.FromArgb(255, 119, 119, 119);
                        if (accuHeight - windowsHeight / 2 > 0)
                            scrollViewer.ScrollToVerticalOffset(accuHeight - windowsHeight / 2);
                        else
                            scrollViewer.ScrollToVerticalOffset(0);
                    }
                    else
                    {
                        if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1 &&
                            docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1 &&
                            File.Exists(Path.Join(compareResultFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png")))
                        {
                            (DocCompareSideListViewLeft.Items[i] as SideGridItemLeft).BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            (DocCompareSideListViewRight.Items[i] as SideGridItemRight).BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                        }
                        else
                        {
                            (DocCompareSideListViewLeft.Items[i] as SideGridItemLeft).BackgroundBrush = Color.FromArgb(0, 255, 255, 255);
                            (DocCompareSideListViewRight.Items[i] as SideGridItemRight).BackgroundBrush = Color.FromArgb(0, 255, 255, 255);
                        }
                    }

                    ListViewItem container = DocCompareSideListViewLeft.ItemContainerGenerator.ContainerFromItem(DocCompareSideListViewLeft.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;
                }
            }
            catch
            {
            }
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            ProcessStartInfo info = new ProcessStartInfo(e.Uri.AbsoluteUri)
            {
                UseShellExecute = true
            };
            Process.Start(info);
            e.Handled = true;
        }

        private void LicenseKeyTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                string currText = (sender as TextBox).Text;
                currText = currText.ToUpper();

                if (licKeyLastInputString == null)
                    licKeyLastInputString = currText;

                if (licKeyLastInputString.Length >= currText.Length)
                {
                    if (currText.Length == 4 || currText.Length == 9 || currText.Length == 14)
                        currText = currText.Remove(currText.Length - 1);
                }
                else
                {
                    if (currText.Length == 4 || currText.Length == 9 || currText.Length == 14)
                        currText += '-';
                }

                if (currText.Length > 19)
                {
                    currText = currText.Remove(currText.Length - 1);
                }

                (sender as TextBox).Text = currText;
                (sender as TextBox).Select(currText.Length, 0);

                if (currText.Length == 19)
                {
                    ValidKeyTick.Visibility = Visibility.Visible;
                    InvalidKeyTick.Visibility = Visibility.Hidden;
                    if (Helper.IsValidEmail(UserEmailTextBox.Text) == true)
                        ActivateLicenseButton.IsEnabled = true;
                    else
                        ActivateLicenseButton.IsEnabled = false;
                }
                else
                {
                    ValidKeyTick.Visibility = Visibility.Hidden;
                    InvalidKeyTick.Visibility = Visibility.Visible;
                    ActivateLicenseButton.IsEnabled = false;
                }

                if (currText.Length == 0)
                {
                    ValidKeyTick.Visibility = Visibility.Hidden;
                    InvalidKeyTick.Visibility = Visibility.Hidden;
                    ActivateLicenseButton.IsEnabled = false;
                }

                licKeyLastInputString = currText;
            }
            catch (Exception ex)
            {
                ErrorHandling.ReportException(ex);
            }
        }

        private void LoadFilesCommonPart()
        {
            SidePanelDocCompareButton.IsEnabled = false;
            /*
            if (settings.numPanelsDragDrop == 3)
                docs.documentsToShow = new List<int>() { 0, 1, 2 };
            else
                docs.documentsToShow = new List<int>() { 0, 1 };
            */

            if (docs.documents.Count >= 1)
            {
                if (docs.documents[docs.documentsToShow[0]].processed == false)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc1Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc1.Visibility = Visibility.Visible;
                        ShowDragDropZone2();
                    });
                }
            }

            if (docs.documents.Count >= 2)
            {
                if (docs.documents[docs.documentsToShow[1]].processed == false)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                {
                    Doc2Grid.Visibility = Visibility.Hidden;
                    ProgressBarDoc2.Visibility = Visibility.Visible;
                    Doc3Grid.Visibility = Visibility.Hidden;
                    DocCompareDragDropZone3.Visibility = Visibility.Visible;
                    ShowDragDropZone3();
                });
                }
            }

            if (docs.documents.Count >= 3)
            {
                if (docs.documents[docs.documentsToShow[2]].processed == false)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc3Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc3.Visibility = Visibility.Visible;
                        Doc4Grid.Visibility = Visibility.Hidden;
                        DocCompareDragDropZone4.Visibility = Visibility.Visible;
                        ShowDragDropZone4();
                    });
                }
            }

            if (docs.documents.Count >= 4)
            {
                if (docs.documents[docs.documentsToShow[3]].processed == false)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc4Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc4.Visibility = Visibility.Visible;
                        Doc5Grid.Visibility = Visibility.Hidden;
                        DocCompareDragDropZone5.Visibility = Visibility.Visible;
                        ShowDragDropZone5();
                    });
                }
            }

            if (docs.documents.Count >= 5)
            {
                if (docs.documents[docs.documentsToShow[4]].processed == false)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc5Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc5.Visibility = Visibility.Visible;
                    });
                }
            }

            docProcessRunning = true;

            Application.Current.Dispatcher.Invoke(() =>
            {
                ProcessingDocProgressCard.Visibility = Visibility.Visible;
                ProcessingDocProgressbar.Value = 0;
                ProcessingDocLabel.Text = "Processing: " + Path.GetFileName(docs.documents[0].filePath);
                /*
                BrowseFileTopButton1.IsEnabled = false;
                BrowseFileTopButton2.IsEnabled = false;
                BrowseFileTopButton3.IsEnabled = false;
                BrowseFileTopButton4.IsEnabled = false;
                BrowseFileTopButton5.IsEnabled = false;
                */
                ReloadDoc1Button.IsEnabled = false;
                ReloadDoc2Button.IsEnabled = false;
                ReloadDoc3Button.IsEnabled = false;
                ReloadDoc4Button.IsEnabled = false;
                ReloadDoc5Button.IsEnabled = false;

                CloseDoc1Button.IsEnabled = false;
                CloseDoc2Button.IsEnabled = false;
                CloseDoc3Button.IsEnabled = false;
                CloseDoc4Button.IsEnabled = false;
                CloseDoc5Button.IsEnabled = false;

                OpenDoc1OriginalButton1.IsEnabled = false;
                OpenDoc2OriginalButton2.IsEnabled = false;
                OpenDoc3OriginalButton3.IsEnabled = false;
                OpenDoc4OriginalButton4.IsEnabled = false;
                OpenDoc5OriginalButton5.IsEnabled = false;

                ShowDoc1FileInfoButton.IsEnabled = false;
                ShowDoc2FileInfoButton.IsEnabled = false;
                ShowDoc3FileInfoButton.IsEnabled = false;
                ShowDoc4FileInfoButton.IsEnabled = false;
                ShowDoc5FileInfoButton.IsEnabled = false;
            });
        }

        private void LoadLicense()
        {
            using var file = File.OpenRead(Path.Join(Path.Join(appDataDir, "lic"), "2compare.lic"));
            lic = Serializer.Deserialize<LicenseManagement>(file);
        }

        private void LoadSettings()
        {
            settings = new AppSettings();
            using var file = File.OpenRead(Path.Join(appDataDir, "AppSettings.bin"));
            settings = Serializer.Deserialize<AppSettings>(file);
        }

        private void MaskSideGridInForceAlignMode()
        {
            string[] splittedNameRef;
            string[] splittedNameTarget;

            if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
            {
                splittedNameRef = selectedSideGridButtonName1.Split("Left");

                foreach (SideGridItemLeft item in DocCompareSideListViewLeft.Items)
                {
                    splittedNameTarget = item.GridName.Split("Grid");
                    if (splittedNameTarget[1] != splittedNameRef[1])
                    {
                        item.GridEffect = new BlurEffect()
                        {
                            Radius = 5,
                        };
                    }
                    else
                    {
                        item.GridEffect = null;
                    }
                }
            }
            else
            {
                splittedNameRef = selectedSideGridButtonName1.Split("Right");

                foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
                {
                    splittedNameTarget = item.GridName.Split("Grid");
                    if (splittedNameTarget[1] != splittedNameRef[1])
                    {
                        item.GridEffect = new BlurEffect()
                        {
                            Radius = 5,
                        };
                    }
                    else
                    {
                        item.GridEffect = null;
                    }
                }
            }
        }

        private void OpenDoc1OriginalButton_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToShow[0]].filePath + "\"";
            fileopener.Start();
        }

        private void OpenDoc2OriginalButton_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToShow[1]].filePath + "\"";
            fileopener.Start();
        }

        private void ProcessDocProgressThread()
        {
            try
            {
                while (docProcessRunning == true)
                {
                    Dispatcher.Invoke(() =>
                    {
                        /*
                        BrowseFileTopButton1.IsEnabled = false;
                        BrowseFileTopButton2.IsEnabled = false;
                        BrowseFileTopButton3.IsEnabled = false;
                        BrowseFileTopButton4.IsEnabled = false;
                        BrowseFileTopButton5.IsEnabled = false;
                        */
                        BrowseFileButton1.IsEnabled = false;
                        BrowseFileButton2.IsEnabled = false;
                        BrowseFileButton3.IsEnabled = false;
                        BrowseFileButton4.IsEnabled = false;
                        BrowseFileButton5.IsEnabled = false;

                        DocCompareFirstDocZone.AllowDrop = false;
                        DocCompareDragDropZone1.AllowDrop = false;
                        DocCompareColorZone1.AllowDrop = false;
                        DocCompareSecondDocZone.AllowDrop = false;
                        DocCompareDragDropZone2.AllowDrop = false;
                        DocCompareColorZone2.AllowDrop = false;
                        DocCompareThirdDocZone.AllowDrop = false;
                        DocCompareDragDropZone3.AllowDrop = false;
                        DocCompareColorZone3.AllowDrop = false;
                        DocCompareFourthDocZone.AllowDrop = false;
                        DocCompareDragDropZone4.AllowDrop = false;
                        DocCompareColorZone4.AllowDrop = false;
                        DocCompareFifthDocZone.AllowDrop = false;
                        DocCompareDragDropZone5.AllowDrop = false;
                        DocCompareColorZone5.AllowDrop = false;

                        Doc1StatsGrid.Visibility = Visibility.Collapsed;
                        Doc2StatsGrid.Visibility = Visibility.Collapsed;
                        Doc3StatsGrid.Visibility = Visibility.Collapsed;
                        Doc4StatsGrid.Visibility = Visibility.Collapsed;
                        Doc5StatsGrid.Visibility = Visibility.Collapsed;

                        //Doc1PageNumberLabel.Content = "";
                        //Doc2PageNumberLabel.Content = "";
                        //Doc3PageNumberLabel.Content = "";
                        //Doc4PageNumberLabel.Content = "";
                        //Doc5PageNumberLabel.Content = "";

                        try
                        {
                            if (docs.documents.Count == 1)
                                ProcessingDocProgressbar.Value = 1;
                            else
                                ProcessingDocProgressbar.Value = ((double)docProcessingCounter / (double)(docs.documents.Count - 1)) * 100.0;

                            ProcessingDocLabel.Text = "Processing: " + Path.GetFileName(docs.documents[docProcessingCounter].filePath);
                        }
                        catch
                        {
                        }
                    });

                    Thread.Sleep(10);
                }

                Dispatcher.Invoke(() =>
                {
                    List<string> idToRemove = new List<string>();
                    for (int i = 0; i < docs.documents.Count; i++)
                    {
                        if (docs.documents[i].processed == false)
                        {
                            idToRemove.Add(docs.documents[i].docID);
                        }
                    }

                    foreach (string id in idToRemove)
                    {
                        docs.RemoveDocumentWithID(id);
                    }

                    ProcessingDocProgressCard.Visibility = Visibility.Hidden;

                    if (docs.documents.Count != 0)
                    {
                        if (docs.documents[docs.documentsToShow[0]].processed == true)
                        {
                            DisplayPreview(1, docs.documentsToShow[0]);
                            ShowDoc1FileInfoButton.IsEnabled = true;
                            ShowDoc2FileInfoButton.IsEnabled = true;
                            OpenDoc1OriginalButton1.IsEnabled = true;
                            DocPreviewStatGrid.Visibility = Visibility.Visible;
                            Doc1StatsGrid.Visibility = Visibility.Visible;
                            UpdateFileStat(0);
                        }
                        /*else
                        {
                            if (docs.documents.Count > 1)
                            {
                                docs.documentsToShow[0] = FindNextDocToShow();
                                DisplayPreview(1, docs.documentsToShow[0]);
                                OpenDoc1OriginalButton1.IsEnabled = true;
                                ShowDoc2FileInfoButton.IsEnabled = true;
                                ShowDoc1FileInfoButton.IsEnabled = true;
                                DocPreviewStatGrid.Visibility = Visibility.Visible;
                                Doc1StatsGrid.Visibility = Visibility.Visible;
                            }
                        }
                        */
                        if (docs.documents.Count > 1)
                        {
                            if (docs.documents[docs.documentsToShow[1]].processed == true)
                            {
                                DisplayPreview(2, docs.documentsToShow[1]);
                                OpenDoc2OriginalButton2.IsEnabled = true;
                                ShowDoc2FileInfoButton.IsEnabled = true;
                                Doc2StatsGrid.Visibility = Visibility.Visible;
                                Doc2NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[1]].filePath);
                                UpdateFileStat(1);
                            }
                            /*
                            else
                            {
                                if (docs.documents.Count > 2)
                                {
                                    docs.documentsToShow[1] = FindNextDocToShow();
                                    DisplayPreview(2, docs.documentsToShow[1]);
                                    OpenDoc2OriginalButton2.IsEnabled = true;
                                    ShowDoc2FileInfoButton.IsEnabled = true;
                                    Doc2StatsGrid.Visibility = Visibility.Visible;
                                }
                            }
                            */
                        }
                        /*
                        else
                        {
                            Doc2StatsGrid.Visibility = Visibility.Collapsed;
                        }
                        */

                        if (docs.documents.Count > 2)
                        {
                            // TODO: Premium
                            //if (settings.numPanelsDragDrop == 3)
                            {
                                if (docs.documents[docs.documentsToShow[2]].processed == true)
                                {
                                    DisplayPreview(3, docs.documentsToShow[2]);
                                    OpenDoc3OriginalButton3.IsEnabled = true;
                                    ShowDoc3FileInfoButton.IsEnabled = true;
                                    Doc3StatsGrid.Visibility = Visibility.Visible;
                                    Doc3NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[2]].filePath);
                                    UpdateFileStat(2);
                                }
                                /*
                                else
                                {
                                    if (docs.documents.Count > 3)
                                    {
                                        docs.documentsToShow[2] = FindNextDocToShow();
                                        DisplayPreview(3, docs.documentsToShow[2]);
                                        OpenDoc3OriginalButton3.IsEnabled = true;
                                    }
                                }
                                */
                            }
                        }

                        if (docs.documents.Count > 3)
                        {
                            // TODO: Premium
                            //if (settings.numPanelsDragDrop == 3)
                            {
                                if (docs.documents[docs.documentsToShow[3]].processed == true)
                                {
                                    DisplayPreview(4, docs.documentsToShow[3]);
                                    OpenDoc4OriginalButton4.IsEnabled = true;
                                    ShowDoc4FileInfoButton.IsEnabled = true;
                                    Doc4StatsGrid.Visibility = Visibility.Visible;
                                    Doc4NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[3]].filePath);
                                    UpdateFileStat(3);
                                }
                                /*
                                else
                                {
                                    if (docs.documents.Count > 4)
                                    {
                                        docs.documentsToShow[3] = FindNextDocToShow();
                                        DisplayPreview(4, docs.documentsToShow[3]);
                                        OpenDoc3OriginalButton3.IsEnabled = true;
                                    }
                                }
                                */
                            }
                        }

                        if (docs.documents.Count > 4)
                        {
                            // TODO: Premium
                            //if (settings.numPanelsDragDrop == 3)
                            {
                                if (docs.documents[docs.documentsToShow[4]].processed == true)
                                {
                                    DisplayPreview(5, docs.documentsToShow[4]);
                                    OpenDoc5OriginalButton5.IsEnabled = true;
                                    ShowDoc5FileInfoButton.IsEnabled = true;
                                    Doc5StatsGrid.Visibility = Visibility.Visible;
                                    Doc5NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[4]].filePath);
                                    UpdateFileStat(4);
                                }
                                /*
                                else
                                {
                                    if (docs.documents.Count > 4)
                                    {
                                        docs.documentsToShow[3] = FindNextDocToShow();
                                        DisplayPreview(4, docs.documentsToShow[3]);
                                        OpenDoc3OriginalButton3.IsEnabled = true;
                                    }
                                }
                                */
                            }
                        }
                    }
                    else
                    {
                        HideDragDropZone2();
                        HideDragDropZone3();
                        HideDragDropZone4();
                        HideDragDropZone5();
                        DocPreviewStatGrid.Visibility = Visibility.Collapsed;
                        Doc1StatsGrid.Visibility = Visibility.Visible;
                        Doc2StatsGrid.Visibility = Visibility.Collapsed;
                        Doc3StatsGrid.Visibility = Visibility.Collapsed;
                        Doc4StatsGrid.Visibility = Visibility.Collapsed;
                        Doc5StatsGrid.Visibility = Visibility.Collapsed;
                    }

                    ProgressBarDoc1.Visibility = Visibility.Hidden;
                    ProgressBarDoc2.Visibility = Visibility.Hidden;
                    ProgressBarDoc3.Visibility = Visibility.Hidden;
                    ProgressBarDoc4.Visibility = Visibility.Hidden;
                    ProgressBarDoc5.Visibility = Visibility.Hidden;
                    /*
                    BrowseFileTopButton1.IsEnabled = true;
                    BrowseFileTopButton2.IsEnabled = true;
                    BrowseFileTopButton3.IsEnabled = true;
                    BrowseFileTopButton4.IsEnabled = true;
                    BrowseFileTopButton5.IsEnabled = true;
                    */
                    ReloadDoc1Button.IsEnabled = true;
                    ReloadDoc2Button.IsEnabled = true;
                    ReloadDoc3Button.IsEnabled = true;
                    ReloadDoc4Button.IsEnabled = true;
                    ReloadDoc5Button.IsEnabled = true;
                    BrowseFileButton1.IsEnabled = true;
                    BrowseFileButton2.IsEnabled = true;
                    BrowseFileButton3.IsEnabled = true;
                    BrowseFileButton4.IsEnabled = true;
                    BrowseFileButton5.IsEnabled = true;

                    DocCompareFirstDocZone.AllowDrop = true;
                    DocCompareDragDropZone1.AllowDrop = true;
                    DocCompareColorZone1.AllowDrop = true;
                    DocCompareSecondDocZone.AllowDrop = true;
                    DocCompareDragDropZone2.AllowDrop = true;
                    DocCompareColorZone2.AllowDrop = true;
                    DocCompareThirdDocZone.AllowDrop = true;
                    DocCompareDragDropZone3.AllowDrop = true;
                    DocCompareColorZone3.AllowDrop = true;
                    DocCompareFourthDocZone.AllowDrop = true;
                    DocCompareDragDropZone4.AllowDrop = true;
                    DocCompareColorZone4.AllowDrop = true;
                    DocCompareFifthDocZone.AllowDrop = true;
                    DocCompareDragDropZone5.AllowDrop = true;
                    DocCompareColorZone5.AllowDrop = true;

                    CloseDoc1Button.IsEnabled = true;
                    CloseDoc2Button.IsEnabled = true;
                    CloseDoc3Button.IsEnabled = true;
                    CloseDoc4Button.IsEnabled = true;
                    CloseDoc5Button.IsEnabled = true;

                    OpenDoc1OriginalButton1.IsEnabled = true;
                    OpenDoc2OriginalButton2.IsEnabled = true;
                    OpenDoc3OriginalButton3.IsEnabled = true;
                    OpenDoc4OriginalButton4.IsEnabled = true;
                    OpenDoc5OriginalButton5.IsEnabled = true;

                    UpdateDocSelectionComboBox();
                    ShowInfoButtonSetVisi();

                    if (docs.documents.Count >= 2)
                        SidePanelDocCompareButton.IsEnabled = true;

                    // Walkthrough
                    if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEBUTTON2 && docs.documents.Count == 3)
                    {
                        PopupDocPreviewNameComboboxBubble.IsOpen = true;
                        PopupBrowseFileButtonBubble.IsOpen = false;
                        PopupBrowseFileButton2Bubble.IsOpen = false;
                        walkthroughStep = WalkthroughSteps.BROWSEFILECOMBOBOX;
                    }
                    else if (walkthroughMode == true)
                    {
                        PopupBrowseFileButtonBubble.IsOpen = false;
                        PopupBrowseFileButton2Bubble.IsOpen = true;
                    }
                    else
                    {
                        PopupBrowseFileButtonBubble.IsOpen = false;
                        PopupBrowseFileButton2Bubble.IsOpen = false;
                    }
                });
            }
            catch (Exception ex)
            {
                ErrorHandling.ReportException(ex);
            }
        }

        private void ShowInfoButtonSetVisi()
        {
            Dispatcher.Invoke(() =>
            {

                //TODO: Premium
                switch (docs.documents.Count)
                {
                    case 1:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Visible;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        break;
                    case 2:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        break;
                    case 3:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Visible;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        break;
                    case 4:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Visible;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        break;
                    case 5:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Visible;
                        break;
                    default:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        break;

                }

            });
        }

        private void ProcessDocThread()
        {
            // Going through documents in stack, check if reloading needed
            docProcessingCounter = 0;

            for (int i = 0; i < docs.documents.Count; i++)
            {
                try
                {
                    if (docs.documents[i].loaded == false && docs.documents[i].filePath != null)
                    {
                        docs.documents[i].loaded = true;
                        docs.documents[i].DetectFileType();
                        docs.documents[i].ReadStats(settings.cultureInfo);

                        int ret = -1;
                        switch (docs.documents[i].fileType)
                        {
                            case Document.FileTypes.PDF:
                                ret = docs.documents[i].ReadPDF();
                                break;

                            case Document.FileTypes.PPT:
                                ret = docs.documents[i].ReadPPT();
                                break;

                            case Document.FileTypes.PIC:
                                ret = docs.documents[i].ReadPic();
                                break;
                        }

                        if (ret == -1)
                        {
                            if (docs.documents[i].fileType == Document.FileTypes.PDF)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    CustomMessageBox msgBox = new CustomMessageBox();
                                    msgBox.Setup("PDF File Corruption", "There was an error converting " + Path.GetFileName(docs.documents[i].filePath) + ". Please repair document and retry.", "Okay");
                                    msgBox.ShowDialog();
                                });
                            }
                            else
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    CustomMessageBox msgBox = new CustomMessageBox();
                                    msgBox.Setup("PowerPoint File Corruption", "There was an error converting " + Path.GetFileName(docs.documents[i].filePath) + ". Please repair document and retry.", "Okay");
                                    msgBox.ShowDialog();
                                });
                            }
                        }
                        else if (ret == -2)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                CustomMessageBox msgBox = new CustomMessageBox();
                                msgBox.Setup("Microsoft PowerPoint not found", "There was an error converting " + Path.GetFileName(docs.documents[i].filePath) + ". No Microsoft PowerPoint installation found.", "Okay");
                                msgBox.ShowDialog();
                            });
                        }
                        else
                        {
                            docs.documents[i].processed = true;
                        }

                        docProcessingCounter += 1;
                    }
                    else
                    {
                        docProcessingCounter += 1;
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandling.ReportException(ex);
                }
            }

            docProcessRunning = false;
        }

        private void RefDocListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int ind = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == RefDocListBox.SelectedItem.ToString());
                DisplayRefDoc(ind);
            }
            catch
            {
            }
        }

        private void RefDocProceedButton_Click(object sender, RoutedEventArgs e)
        {
            docs.documentsToCompare[0] = RefDocListBox.SelectedIndex;
            UpdateDocCompareComboBox();
            SetVisiblePanel(SidePanels.DOCCOMPARE);
        }

        private void RefDocScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                double accuHeight = 0;

                for (int i = 0; i < refDocPanel.Children.Count; i++)
                {
                    Size currSize = refDocPanel.Children[i].DesiredSize;
                    accuHeight += currSize.Height;

                    if (accuHeight > RefDocScrollViewer.VerticalOffset)
                    {
                        RefDocPageNumberLabel.Content = (i + 1).ToString() + " / " + refDocPanel.Children.Count.ToString();
                        break;
                    }
                }
            }
            catch
            {
            }
        }

        private void ReloadDoc1Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDoc1.Visibility = Visibility.Visible;

            docs.docToReload = docs.documentsToShow[0];
            docs.displayToReload = 0;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDoc2Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDoc2.Visibility = Visibility.Visible;

            docs.docToReload = docs.documentsToShow[1];
            docs.displayToReload = 1;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDoc3Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDoc3.Visibility = Visibility.Visible;

            docs.docToReload = docs.documentsToShow[2];
            docs.displayToReload = 2;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDocCompare1Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDocCompareReload.Visibility = Visibility.Visible;
            docs.forceAlignmentIndices = new List<List<int>>();

            docs.docToReload = docs.documentsToCompare[0];
            docs.displayToReload = 5;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDocCompare2Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDocCompareReload.Visibility = Visibility.Visible;
            docs.forceAlignmentIndices = new List<List<int>>();
            docs.docToReload = docs.documentsToCompare[1];
            docs.displayToReload = 6;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();

            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPARERELOAD)
            {
                PopupReloadBubble.IsOpen = false;
                walkthroughStep = WalkthroughSteps.COMPARELINK; // 11
            }
        }

        private void ReloadDocThread()
        {
            Dispatcher.Invoke(() =>
            {
                ReloadDoc1Button.IsEnabled = false;
                ReloadDoc2Button.IsEnabled = false;
                ReloadDoc3Button.IsEnabled = false;
                ReloadDocCompare1Button.IsEnabled = false;
                ReloadDocCompare2Button.IsEnabled = false;
            });

            if (docs.documents[docs.docToReload].ReloadDocument(workingDir) == 0)
            {
                docs.documents[docs.docToReload].ReadStats(settings.cultureInfo);

                Dispatcher.Invoke(() =>
                {
                    switch (docs.displayToReload)
                    {
                        case 0:
                            DisplayPreview(1, docs.documentsToShow[0]);
                            //DisplayImageLeft(docs.documentsToShow[0]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(0);
                            break;

                        case 1:
                            DisplayPreview(2, docs.documentsToShow[1]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(1);
                            break;

                        case 2:
                            DisplayPreview(3, docs.documentsToShow[2]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(2);
                            break;

                        case 3:
                            DisplayPreview(4, docs.documentsToShow[3]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(3);
                            break;

                        case 4:
                            DisplayPreview(5, docs.documentsToShow[4]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(4);
                            break;

                        case 5:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[0])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayPreview(1, docs.documentsToShow[0]);
                                            //DisplayImageLeft(docs.documentsToShow[0]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(1);
                                            break;

                                        case 1:
                                            DisplayPreview(2, docs.documentsToShow[1]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(2);
                                            break;

                                        case 2:
                                            DisplayPreview(3, docs.documentsToShow[2]);
                                            break;

                                        case 3:
                                            DisplayPreview(4, docs.documentsToShow[3]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(3);
                                            break;

                                        case 4:
                                            DisplayPreview(5, docs.documentsToShow[4]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(4);
                                            break;
                                    }
                                }
                            }

                            ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                            docCompareGrid.Visibility = Visibility.Hidden;
                            docCompareSideGridShown = 0;
                            DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                            DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                            DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                            SetVisiblePanel(SidePanels.DOCCOMPARE);
                            ProgressBarDocCompare.Visibility = Visibility.Visible;
                            threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                            threadCompare.Start();
                            break;

                        case 6:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[1])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayPreview(1, docs.documentsToShow[0]);
                                            //DisplayImageLeft(docs.documentsToShow[0]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(0);
                                            break;

                                        case 1:
                                            DisplayPreview(2, docs.documentsToShow[1]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(1);
                                            break;

                                        case 2:
                                            DisplayPreview(3, docs.documentsToShow[2]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(2);
                                            break;

                                        case 3:
                                            DisplayPreview(4, docs.documentsToShow[3]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(3);
                                            break;

                                        case 4:
                                            DisplayPreview(5, docs.documentsToShow[4]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(4);
                                            break;
                                    }
                                }
                            }

                            ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                            docCompareGrid.Visibility = Visibility.Hidden;
                            docCompareSideGridShown = 0;
                            DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                            DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                            DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                            SetVisiblePanel(SidePanels.DOCCOMPARE);
                            ProgressBarDocCompare.Visibility = Visibility.Visible;
                            threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                            threadCompare.Start();
                            break;
                    }

                    ReloadDoc1Button.IsEnabled = true;
                    ReloadDoc2Button.IsEnabled = true;
                    ReloadDoc3Button.IsEnabled = true;
                    ReloadDocCompare1Button.IsEnabled = true;
                    ReloadDocCompare2Button.IsEnabled = true;
                });
            }
            else
            {
                docs.documents[docs.docToReload].processed = false;

                Dispatcher.Invoke(() =>
                {
                    switch (docs.displayToReload)
                    {
                        case 0:
                            DisplayPreview(1, docs.documentsToShow[0]);
                            //DisplayImageLeft(docs.documentsToShow[0]);
                            break;

                        case 1:
                            DisplayPreview(2, docs.documentsToShow[1]);
                            break;

                        case 2:
                            DisplayPreview(3, docs.documentsToShow[2]);
                            break;

                        case 3:
                            DisplayPreview(4, docs.documentsToShow[3]);
                            break;

                        case 4:
                            DisplayPreview(5, docs.documentsToShow[4]);
                            break;

                        case 5:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[0])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayPreview(1, docs.documentsToShow[0]);
                                            //DisplayImageLeft(docs.documentsToShow[0]);
                                            break;

                                        case 1:
                                            DisplayPreview(2, docs.documentsToShow[1]);
                                            break;

                                        case 2:
                                            DisplayPreview(3, docs.documentsToShow[2]);
                                            break;

                                        case 3:
                                            DisplayPreview(4, docs.documentsToShow[3]);
                                            break;

                                        case 4:
                                            DisplayPreview(5, docs.documentsToShow[4]);
                                            break;
                                    }
                                }
                            }

                            ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                            break;

                        case 6:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[1])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayPreview(1, docs.documentsToShow[0]);
                                            //DisplayImageLeft(docs.documentsToShow[0]);
                                            break;

                                        case 1:
                                            DisplayPreview(2, docs.documentsToShow[1]);
                                            break;

                                        case 2:
                                            DisplayPreview(3, docs.documentsToShow[2]);
                                            break;

                                        case 3:
                                            DisplayPreview(4, docs.documentsToShow[3]);
                                            break;

                                        case 4:
                                            DisplayPreview(5, docs.documentsToShow[4]);
                                            break;
                                    }
                                }
                            }

                            ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                            break;
                    }
                });

                if (docs.documents[docs.docToReload].fileType == Document.FileTypes.PDF)
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("PDF File Corruption", "There was an error converting " + Path.GetFileName(docs.documents[docs.docToReload].filePath) + ". Please repair document and retry.", "Okay");
                    msgBox.ShowDialog();
                }
                else
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("PowerPoint File Corruption", "There was an error converting " + Path.GetFileName(docs.documents[docs.docToReload].filePath) + ". Please repair document and retry.", "Okay");
                    msgBox.ShowDialog();
                }
            }
        }

        private void RemoveAllForceAlignButton_Click(object sender, RoutedEventArgs e)
        {
            if (docs.forceAlignmentIndices.Count != 0)
            {
                docs.forceAlignmentIndices = new List<List<int>>();
                inForceAlignMode = false;

                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;

                DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                ProgressBarDocCompareAlign.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();

                Dispatcher.Invoke(() =>
                {
                    UnMaskSideGridFromForceAlignMode();
                    EnableRemoveForceAlignButton();
                    EnableSideScrollLeft();
                    EnableSideScrollRight();
                });
            }

            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPAREUNLINK)
            {
                PopupUnlinkBubble.IsOpen = false;
                ListViewItem container = DocCompareMainListView.ItemContainerGenerator.ContainerFromItem(DocCompareMainListView.Items[0]) as ListViewItem;
                Grid thisGrid = VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(container, 0) as Grid, 1) as Grid;
                Button thisButton = VisualTreeHelper.GetChild(thisGrid, 3) as Button;
                thisButton.Visibility = Visibility.Visible;

                PopupAnimateBubble.PlacementTarget = container;

                PopupAnimateBubble.IsOpen = true;
                walkthroughStep = WalkthroughSteps.COMPAREANIMATE;
            }
        }

        private void RemoveForceAlignClicked(object sender, RoutedEventArgs args)
        {
            if (inForceAlignMode == false)
            {
                Button button = sender as Button;
                string[] splittedName = button.Tag.ToString().Split("Align");
                docs.RemoveForceAligmentPairs(docs.documents[docs.documentsToCompare[0]].docCompareIndices[int.Parse(splittedName[1])]);

                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
        }

        private async void RenewLicense()
        {
            LicenseManagement.LicServerResponse res = await lic.RenewLicense();
            CustomMessageBox msgBox;
            Dispatcher.Invoke(() =>
            {
                switch (res)
                {
                    case LicenseManagement.LicServerResponse.UNREACHABLE:
                        TimeSpan bufferTime = lic.GetExpiryWaiveDate().Subtract(DateTime.Today);
                        if (bufferTime.TotalDays >= 0)
                        {
                            msgBox = new CustomMessageBox();
                            msgBox.Setup("No connection to license server", "The 2|Compare license server cannot be reached. This may be due to multiple reasons, including your firewall settings or no connection to the Internet but also server maintenance on our side. If you are connected to the Internet and your firewall is configured properly, please ignore this warning. Out server will be up and running again shortly. If you see this warning throughout multiple consecutive days, please contact us at support@hopie.tech.", "Okay");
                            msgBox.ShowDialog();
                        }
                        else
                        {
                            msgBox = new CustomMessageBox();
                            msgBox.Setup("No connection to license server", "The 2|Compare license server cannot be reached. This may be due to multiple reasons, including your firewall settings or no connection to the Internet but also server maintenance on our side. If you are connected to the Internet and your firewall is configured properly, please ignore this warning. Out server will be up and running again shortly. If you see this warning throughout multiple consecutive days, please contact us at support@hopie.tech.", "Okay");
                            msgBox.ShowDialog();
                            BrowseFileButton1.IsEnabled = false;
                            DocCompareFirstDocZone.AllowDrop = false;
                            DocCompareDragDropZone1.AllowDrop = false;
                            DocCompareColorZone1.AllowDrop = false;
                        }
                        UserEmailTextBox.IsEnabled = true;
                        LicenseKeyTextBox.IsEnabled = true; // after successful activation, we will prevent further editing
                        ActivateLicenseButton.IsEnabled = true;
                        SaveLicense();
                        break;

                    case LicenseManagement.LicServerResponse.KEY_MISMATCH:
                        msgBox = new CustomMessageBox();
                        msgBox.Setup("Invalid license key", "The provided license key does not match the email address. Please check your inputs.", "Okay");
                        msgBox.ShowDialog();
                        UserEmailTextBox.IsEnabled = true;
                        LicenseKeyTextBox.IsEnabled = true; // after successful activation, we will prevent further editing
                        ActivateLicenseButton.IsEnabled = true;
                        break;

                    case LicenseManagement.LicServerResponse.ACCOUNT_NOT_FOUND:
                        msgBox = new CustomMessageBox();
                        msgBox.Setup("License not found", "No license was found under the given email address. Please check your inputs.", "Okay");
                        msgBox.ShowDialog();
                        UserEmailTextBox.IsEnabled = true;
                        LicenseKeyTextBox.IsEnabled = true; // after successful activation, we will prevent further editing
                        ActivateLicenseButton.IsEnabled = true;
                        break;

                    case LicenseManagement.LicServerResponse.OKAY:
                        msgBox = new CustomMessageBox();
                        msgBox.Setup("License renewal", "License renewed successfully.", "Okay");
                        msgBox.ShowDialog();
                        UserEmailTextBox.IsEnabled = false;
                        LicenseKeyTextBox.IsEnabled = false; // after successful activation, we will prevent further editing
                        ActivateLicenseButton.IsEnabled = false;
                        SaveLicense(); // only save license info if successful
                                       // allow usage if it was previously disabled
                        DisplayLicense();
                        break;
                }
            });
        }

        private void SaveLicense()
        {
            Directory.CreateDirectory(Path.Join(appDataDir, "lic"));
            using var file = File.Create(Path.Join(Path.Join(appDataDir, "lic"), "2compare.lic"));
            Serializer.Serialize(file, lic);
        }

        private void SaveSettings()
        {
            using var file = File.Create(Path.Join(appDataDir, "AppSettings.bin"));
            Serializer.Serialize(file, settings);
        }

        private void SettingsAboutButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleSettingsPanel(SettingsPanels.ABOUT);
        }

        private void SettingsBrowDefaultFolderButton_Click(object sender, RoutedEventArgs e)
        {
        }

        private void SettingsBrowseDocButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleSettingsPanel(SettingsPanels.DOCBROWSING);
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisiblePanel(SidePanels.SETTINGS);
            SetVisibleSettingsPanel(SettingsPanels.ABOUT);
        }

        private void SettingsShowThirdPanelCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            settings.numPanelsDragDrop = 3;
            SaveSettings();
            docs.documentsToShow = new List<int>() { 0, 1, 2 };
            if (docs.documents.Count >= 1)
                DisplayPreview(1, docs.documentsToShow[0]);
            //DisplayImageLeft(docs.documentsToShow[0]);
            if (docs.documents.Count >= 2)
                DisplayPreview(2, docs.documentsToShow[1]);
            if (docs.documents.Count >= 2)
            {
                ShowDragDropZone3();
                Doc3Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone3.Visibility = Visibility.Visible;
            }

            if (docs.documents.Count >= 3)
                DisplayPreview(3, docs.documentsToShow[2]);

            UpdateDocSelectionComboBox();
        }

        private void SettingsShowThirdPanelCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            settings.numPanelsDragDrop = 2;
            SaveSettings();
            docs.documentsToShow = new List<int>() { 0, 1 };
            DisplayPreview(1, docs.documentsToShow[0]);
            DisplayPreview(2, docs.documentsToShow[1]);
            //DisplayImageLeft(docs.documentsToShow[0]);
            //DisplayPreview(2, docs.documentsToShow[1]);
            HideDragDropZone3();
            UpdateDocSelectionComboBox();
        }

        private void SettingsSubscriptionButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisibleSettingsPanel(SettingsPanels.SUBSCRIPTION);
        }

        private void SetVisiblePanel(SidePanels p_sidePanel)
        {
            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            currentVisiblePanel = p_sidePanel;

            switch (p_sidePanel)
            {
                case SidePanels.DRAGDROP:
                    SidePanelOpenDocBackground.Background = brush;
                    SidePanelDocCompareBackground.Background = Brushes.Transparent;
                    SettingsButtonBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Visible;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    SettingsPanel.Visibility = Visibility.Hidden;
                    SelectReferenceDocPanel.Visibility = Visibility.Hidden;
                    break;

                case SidePanels.DOCCOMPARE:
                    SidePanelOpenDocBackground.Background = Brushes.Transparent;
                    SidePanelDocCompareBackground.Background = brush;
                    SettingsButtonBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Visible;
                    SettingsPanel.Visibility = Visibility.Hidden;
                    SelectReferenceDocPanel.Visibility = Visibility.Hidden;
                    break;

                case SidePanels.SETTINGS:
                    SidePanelOpenDocBackground.Background = Brushes.Transparent;
                    SidePanelDocCompareBackground.Background = Brushes.Transparent;
                    SettingsButtonBackground.Background = brush;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    SettingsPanel.Visibility = Visibility.Visible;
                    SelectReferenceDocPanel.Visibility = Visibility.Hidden;
                    break;

                case SidePanels.REFDOC:
                    SidePanelOpenDocBackground.Background = Brushes.Transparent;
                    SidePanelDocCompareBackground.Background = brush;
                    SettingsButtonBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    SettingsPanel.Visibility = Visibility.Hidden;
                    SelectReferenceDocPanel.Visibility = Visibility.Visible;
                    break;

                default:
                    SidePanelOpenDocBackground.Background = Brushes.Transparent;
                    SidePanelDocCompareBackground.Background = Brushes.Transparent;
                    SettingsButtonBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    SettingsPanel.Visibility = Visibility.Hidden;
                    SelectReferenceDocPanel.Visibility = Visibility.Hidden;
                    break;
            }
        }

        private void SetVisibleSettingsPanel(SettingsPanels p_settingsPanel)
        {
            Brush brush = FindResource("SidePanelActiveButton") as Brush;

            switch (p_settingsPanel)
            {
                case SettingsPanels.DOCBROWSING:
                    SettingsDocBrowsingPanel.Visibility = Visibility.Visible;
                    SettingsAboutPanel.Visibility = Visibility.Hidden;
                    SettingsSubscriptionPanel.Visibility = Visibility.Hidden;
                    //SettingsBrowseDocButton.Background = brush;
                    //SettingsDocCompareButton.Background = Brushes.Transparent;
                    SettingsSubscriptionButton.Background = Brushes.Transparent;
                    SettingsAboutButton.Background = Brushes.Transparent;
                    break;

                case SettingsPanels.ABOUT:
                    SettingsDocBrowsingPanel.Visibility = Visibility.Hidden;
                    SettingsAboutPanel.Visibility = Visibility.Visible;
                    SettingsSubscriptionPanel.Visibility = Visibility.Hidden;
                    //SettingsBrowseDocButton.Background = Brushes.Transparent;
                    //SettingsDocCompareButton.Background = Brushes.Transparent;
                    SettingsSubscriptionButton.Background = Brushes.Transparent;
                    SettingsAboutButton.Background = brush;
                    break;

                case SettingsPanels.SUBSCRIPTION:
                    SettingsDocBrowsingPanel.Visibility = Visibility.Hidden;
                    SettingsAboutPanel.Visibility = Visibility.Hidden;
                    SettingsSubscriptionPanel.Visibility = Visibility.Visible;
                    //SettingsBrowseDocButton.Background = Brushes.Transparent;
                    //SettingsDocCompareButton.Background = Brushes.Transparent;
                    SettingsSubscriptionButton.Background = brush;
                    SettingsAboutButton.Background = Brushes.Transparent;
                    break;
            }
        }

        private void ShowDoc1FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(0);

            if (Doc1StatAuthorLabel0.Visibility == Visibility.Collapsed)
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc1StatAuthorLabel.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Visible;
            }
            else
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc1StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Collapsed;
            }

            /*
            if(Doc1StatsGrid.ActualHeight > Doc2StatsGrid.ActualHeight)
                Doc2StatsGrid.Height = Doc1StatsGrid.ActualHeight;
            else
                Doc1StatsGrid.Height = Doc2StatsGrid.ActualHeight;
            */
        }

        private void ShowDoc2FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(1);

            if (Doc1StatAuthorLabel0.Visibility == Visibility.Collapsed)
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc1StatAuthorLabel.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Visible;
            }
            else
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc1StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Collapsed;
            }

            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEINFOOPEN)
            {
                PopupDocPreviewInfoButtonBubble.IsOpen = false;
                PopupDocPreviewInfoButton2Bubble.IsOpen = true;
                walkthroughStep = WalkthroughSteps.BROWSEFILEINFOCLOSE; // 5
            }
            else if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILEINFOCLOSE)
            {
                PopupDocPreviewInfoButton2Bubble.IsOpen = false;
                PopupCompareDocBubble.IsOpen = true;
                walkthroughStep = WalkthroughSteps.COMPARETAB; // 6
            }

            /*
            if(Doc1StatsGrid.ActualHeight > Doc2StatsGrid.ActualHeight)
                Doc2StatsGrid.Height = Doc1StatsGrid.ActualHeight;
            else
                Doc1StatsGrid.Height = Doc2StatsGrid.ActualHeight;
            */
        }

        private void ShowDoc3FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(2);

            if (Doc1StatAuthorLabel0.Visibility == Visibility.Collapsed)
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc1StatAuthorLabel.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Visible;
            }
            else
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc1StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Collapsed;
            }
        }

        private void ShowDocCompareFileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(5);

            if (DocCompareLeftStatAuthorLabel.Visibility == Visibility.Collapsed)
            {
                DocCompareLeftStatAuthorLabel.Visibility = Visibility.Visible;
                DocCompareLeftStatCreatedLabel.Visibility = Visibility.Visible;
                DocCompareLeftStatAuthorLabel0.Visibility = Visibility.Visible;
                DocCompareLeftStatCreatedLabel0.Visibility = Visibility.Visible;
                DocCompareLeftStatLastEditorLabel0.Visibility = Visibility.Visible;
                DocCompareLeftStatLastEditorLabel.Visibility = Visibility.Visible;
            }
            else
            {
                DocCompareLeftStatAuthorLabel.Visibility = Visibility.Collapsed;
                DocCompareLeftStatCreatedLabel.Visibility = Visibility.Collapsed;
                DocCompareLeftStatAuthorLabel0.Visibility = Visibility.Collapsed;
                DocCompareLeftStatCreatedLabel0.Visibility = Visibility.Collapsed;
                DocCompareLeftStatLastEditorLabel0.Visibility = Visibility.Collapsed;
                DocCompareLeftStatLastEditorLabel.Visibility = Visibility.Collapsed;
            }

            if (DocCompareRightStatAuthorLabel.Visibility == Visibility.Collapsed)
            {
                DocCompareRightStatAuthorLabel.Visibility = Visibility.Visible;
                DocCompareRightStatCreatedLabel.Visibility = Visibility.Visible;
                DocCompareRightStatAuthorLabel0.Visibility = Visibility.Visible;
                DocCompareRightStatCreatedLabel0.Visibility = Visibility.Visible;
                DocCompareRightStatLastEditorLabel0.Visibility = Visibility.Visible;
                DocCompareRightStatLastEditorLabel.Visibility = Visibility.Visible;
            }
            else
            {
                DocCompareRightStatAuthorLabel.Visibility = Visibility.Collapsed;
                DocCompareRightStatCreatedLabel.Visibility = Visibility.Collapsed;
                DocCompareRightStatCreatedLabel0.Visibility = Visibility.Collapsed;
                DocCompareRightStatAuthorLabel0.Visibility = Visibility.Collapsed;
                DocCompareRightStatAuthorLabel0.Visibility = Visibility.Collapsed;
                DocCompareRightStatLastEditorLabel0.Visibility = Visibility.Collapsed;
                DocCompareRightStatLastEditorLabel.Visibility = Visibility.Collapsed;
            }
        }

        private void ShowDragDropZone2()
        {
            DocCompareSecondDocZone.Visibility = Visibility.Visible;
            DragDropPanel.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Star);
            //Doc2PageNumberLabel.Visibility = Visibility.Visible;
            Doc2StatsGrid.Visibility = Visibility.Visible;
            ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
            DocPreviewStatGrid.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Star);
        }

        private void ShowDragDropZone3()
        {
            DocCompareThirdDocZone.Visibility = Visibility.Visible;
            DragDropPanel.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
            //Doc3PageNumberLabel.Visibility = Visibility.Visible;
            Doc3StatsGrid.Visibility = Visibility.Visible;
            ShowDoc3FileInfoButton.Visibility = Visibility.Visible;
            DocPreviewStatGrid.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
        }

        private void ShowDragDropZone4()
        {
            DocCompareFourthDocZone.Visibility = Visibility.Visible;
            DragDropPanel.ColumnDefinitions[3].Width = new GridLength(1, GridUnitType.Star);
            //Doc4PageNumberLabel.Visibility = Visibility.Visible;
            Doc4StatsGrid.Visibility = Visibility.Visible;
            ShowDoc4FileInfoButton.Visibility = Visibility.Visible;
            DocPreviewStatGrid.ColumnDefinitions[3].Width = new GridLength(1, GridUnitType.Star);
        }

        private void ShowDragDropZone5()
        {
            DocCompareFifthDocZone.Visibility = Visibility.Visible;
            DragDropPanel.ColumnDefinitions[4].Width = new GridLength(1, GridUnitType.Star);
            //Doc5PageNumberLabel.Visibility = Visibility.Visible;
            Doc5StatsGrid.Visibility = Visibility.Visible;
            ShowDoc5FileInfoButton.Visibility = Visibility.Visible;
            DocPreviewStatGrid.ColumnDefinitions[4].Width = new GridLength(1, GridUnitType.Star);
        }
        private void ShowExistingDocCountWarningBox(string docName)
        {
            CustomMessageBox msgBox = new CustomMessageBox();
            msgBox.Setup("Document already loaded", "This document has been loaded: " + docName, "Okay");
            msgBox.ShowDialog();
        }

        private void ShowInvalidDocTypeWarningBox(string fileType, string filename)
        {
            CustomMessageBox msgBox = new CustomMessageBox();
            msgBox.Setup("Unsupported file type", "Unsupported file type of " + fileType + " selected with " + filename + ". This document will be ignored.", "Okay");
            msgBox.ShowDialog();
        }

        private void ShowMaskButton_Click(object sender, RoutedEventArgs e)
        {
            showMask = true;
            /*
            DisplayComparisonResult();
            Border border = (Border)VisualTreeHelper.GetChild(DocCompareMainListView, 0);
            ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

            double currOffset = scrollViewer.VerticalOffset;
            scrollViewer.ScrollToVerticalOffset(0);
            scrollViewer.ScrollToVerticalOffset(currOffset);
            //HighlightSideGrid();
            */
            foreach (CompareMainItem item in DocCompareMainListView.Items)
            {
                item.ShowMask = Visibility.Visible;
            }

            foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
            {
                item.ShowMask = Visibility.Visible;
            }

            ShowMaskButton.Visibility = Visibility.Hidden;
            HideMaskButton.Visibility = Visibility.Visible;
            HighlightingDisableTip.Visibility = Visibility.Hidden;
        }

        private void ShowMaxDocCountWarningBox()
        {
            CustomMessageBox msgBox = new CustomMessageBox();
            msgBox.Setup("Maximum documents loaded", "You have selected more than " + settings.maxDocCount.ToString() + " documents. Only the first " + settings.maxDocCount.ToString() + " documents are loaded.", "Okay");
            msgBox.ShowDialog();
        }

        private void SideGridButtonMouseClick(object sender, RoutedEventArgs args)
        {
            Border border = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewLeft, 0);
            ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;
            Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewRight, 0);
            ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

            Button button = sender as Button;
            if (inForceAlignMode == false)
            {
                scrollPosLeft = scrollViewer.VerticalOffset;
                scrollPosRight = scrollViewer2.VerticalOffset;

                selectedSideGridButtonName1 = button.Tag.ToString();
                if (selectedSideGridButtonName1.Contains("Left"))
                {
                    sideGridSelectedLeftOrRight = GridSelection.LEFT;
                    Dispatcher.Invoke(() => { DisableSideScrollLeft(); });
                }
                else
                {
                    sideGridSelectedLeftOrRight = GridSelection.RIGHT;
                    Dispatcher.Invoke(() => { DisableSideScrollRight(); });
                }

                inForceAlignMode = true;
                Dispatcher.Invoke(() =>
                {
                    MaskSideGridInForceAlignMode();
                    DisableRemoveForceAlignButton();

                    if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPARELINK)
                    {
                        PopupLinkPageBubble.IsOpen = false;
                        walkthroughStep = WalkthroughSteps.COMPAREUNLINK;
                    }
                });
            }
            else
            {
                selectedSideGridButtonName2 = button.Tag.ToString();

                if (selectedSideGridButtonName2 == selectedSideGridButtonName1) // click on original page again
                {
                    Dispatcher.Invoke(() =>
                    {
                        UnMaskSideGridFromForceAlignMode();
                        EnableRemoveForceAlignButton();

                        if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                        {
                            //DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(scrollPosLeft);
                            scrollViewer2.ScrollToVerticalOffset(scrollPosLeft);
                        }
                        else
                        {
                            scrollViewer.ScrollToVerticalOffset(scrollPosRight);
                            //DocCompareSideScrollViewerRight.ScrollToVerticalOffset(scrollPosRight);
                        }
                    });
                    inForceAlignMode = false;
                    return;
                }
                else // another page selected
                {
                    inForceAlignMode = false;

                    int source, target;
                    string[] splittedNameTarget;
                    string[] splittedNameRef;

                    if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                    {
                        splittedNameRef = selectedSideGridButtonName1.Split("Left");
                        splittedNameTarget = selectedSideGridButtonName2.Split("Right");

                        source = docs.documents[docs.documentsToCompare[0]].docCompareIndices[int.Parse(splittedNameRef[1])];
                        target = docs.documents[docs.documentsToCompare[1]].docCompareIndices[int.Parse(splittedNameTarget[1])];
                        docs.AddForceAligmentPairs(source, target);
                    }
                    else
                    {
                        splittedNameRef = selectedSideGridButtonName2.Split("Left");
                        splittedNameTarget = selectedSideGridButtonName1.Split("Right");

                        source = docs.documents[docs.documentsToCompare[0]].docCompareIndices[int.Parse(splittedNameRef[1])];
                        target = docs.documents[docs.documentsToCompare[1]].docCompareIndices[int.Parse(splittedNameTarget[1])];
                        docs.AddForceAligmentPairs(source, target);
                    }

                    ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                    docCompareGrid.Visibility = Visibility.Hidden;
                    docCompareSideGridShown = 0;
                    DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                    DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                    DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                    SetVisiblePanel(SidePanels.DOCCOMPARE);
                    ProgressBarDocCompareAlign.Visibility = Visibility.Visible;
                    threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                    threadCompare.Start();

                    Dispatcher.Invoke(() =>
                    {
                        UnMaskSideGridFromForceAlignMode();
                        EnableRemoveForceAlignButton();
                        EnableSideScrollLeft();
                        EnableSideScrollRight();

                        if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPAREUNLINK)
                        {
                            PopupUnlinkBubble.IsOpen = true;
                        }
                    });
                }
            }
        }

        private void SideGridMouseEnter(object sender, MouseEventArgs args)
        {
            Grid img = sender as Grid;
            string imgName = img.Tag.ToString();
            string[] splittedName;
            string nameToLook = "";
            string nameToLook1 = "";

            if (imgName.Contains("Left"))
            {
                splittedName = imgName.Split("Left");
                nameToLook = "SideButtonLeft" + splittedName[1];
            }
            else
            {
                splittedName = imgName.Split("Right");
                nameToLook = "SideButtonRight" + splittedName[1];
            }

            int selfInd = int.Parse(splittedName[1]);

            int refInd = 0;
            if (inForceAlignMode)
            {
                string[] splittedRefName;
                if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                {
                    splittedRefName = selectedSideGridButtonName1.Split("Left");
                    refInd = int.Parse(splittedRefName[1]);
                }
                else
                {
                    splittedRefName = selectedSideGridButtonName1.Split("Right");
                    refInd = int.Parse(splittedRefName[1]);
                }
            }

            int lowerIndLeft = 0;
            int lowerIndRight = 0;
            int upperIndLeft = docs.totalLen;
            int upperIndRight = docs.totalLen;
            int pairLeft = 0;
            int pairRight = 0;
            bool isLinkedPage = false;

            //if (inForceAlignMode == true)
            {
                foreach (List<int> pair in docs.forceAlignmentIndices)
                {
                    pairLeft = docs.documents[docs.documentsToCompare[0]].docCompareIndices.FindIndex(x => x == pair[0]);
                    pairRight = docs.documents[docs.documentsToCompare[1]].docCompareIndices.FindIndex(x => x == pair[1]);

                    if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                    {
                        if (selfInd == pairRight)
                            isLinkedPage |= true;

                        if (pairLeft > lowerIndLeft && pairLeft <= refInd)
                        {
                            lowerIndLeft = pairLeft;
                            lowerIndRight = pairRight;
                        }

                        if (pairLeft < upperIndLeft && pairLeft >= refInd)
                        {
                            upperIndLeft = pairLeft;
                            upperIndRight = pairRight;
                        }
                    }
                    else
                    {
                        if (selfInd == pairLeft)
                            isLinkedPage |= true;

                        if (pairRight > lowerIndRight && pairRight <= refInd)
                        {
                            lowerIndLeft = pairLeft;
                            lowerIndRight = pairRight;
                        }

                        if (pairRight < upperIndRight && pairRight >= refInd)
                        {
                            upperIndLeft = pairLeft;
                            upperIndRight = pairRight;
                        }
                    }
                }
            }

            if (isLinkedPage == true)
            {
                nameToLook1 = "RemoveForceAlign" + splittedName[1];
            }
            else
            {
                nameToLook1 = nameToLook;
            }

            if (inForceAlignMode == false)
            {
                foreach (object child in img.Children)
                {
                    if (child is Button)
                    {
                        if ((child as Button).Tag.ToString() == nameToLook1)
                        {
                            (child as Button).Visibility = Visibility.Visible;
                        }
                        else
                        {
                            (child as Button).Visibility = Visibility.Hidden;
                        }
                    }
                }
            }
            else
            {
                foreach (object child in img.Children)
                {
                    if (child is Button)
                    {
                        if ((child as Button).Tag.ToString() == nameToLook)
                        {
                            Button foundButton = child as Button;
                            if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                            {
                                if (nameToLook.Contains("Left"))
                                {
                                    if (nameToLook != selectedSideGridButtonName1)
                                        foundButton.Visibility = Visibility.Hidden;
                                    else
                                        foundButton.Visibility = Visibility.Visible;
                                }
                                else
                                {
                                    if (lowerIndRight <= selfInd && selfInd <= upperIndLeft && isLinkedPage == false)
                                        foundButton.Visibility = Visibility.Visible;
                                    else
                                    {
                                        nameToLook = "SideButtonInvalidRight" + splittedName[1];
                                        foreach (object child2 in img.Children)
                                        {
                                            if (child2 is Button)
                                            {
                                                if ((child2 as Button).Tag.ToString() == nameToLook)
                                                {
                                                    foundButton = child2 as Button;

                                                    if (isLinkedPage == true)
                                                    {
                                                        foundButton.ToolTip = "Page linked. Please remove existing link first.";
                                                    }
                                                    else
                                                    {
                                                        foundButton.ToolTip = "This link would cross a previously set link. Please remove that link before aligning these pages.";
                                                    }

                                                    foundButton.Visibility = Visibility.Visible;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (nameToLook.Contains("Right"))
                                {
                                    if (nameToLook != selectedSideGridButtonName1)
                                        foundButton.Visibility = Visibility.Hidden;
                                    else
                                        foundButton.Visibility = Visibility.Visible;
                                }
                                else
                                {
                                    if (lowerIndRight <= selfInd && selfInd <= upperIndRight && isLinkedPage == false)
                                        foundButton.Visibility = Visibility.Visible;
                                    else
                                    {
                                        nameToLook = "SideButtonInvalidLeft" + splittedName[1];
                                        foreach (object child2 in img.Children)
                                        {
                                            if (child2 is Button)
                                            {
                                                if ((child2 as Button).Tag.ToString() == nameToLook)
                                                {
                                                    foundButton = child2 as Button;

                                                    if (isLinkedPage == true)
                                                    {
                                                        foundButton.ToolTip = "Page linked. Please remove existing link first.";
                                                    }
                                                    else
                                                    {
                                                        foundButton.ToolTip = "This link would cross a previously set link. Please remove that link before aligning these pages.";
                                                    }

                                                    foundButton.Visibility = Visibility.Visible;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            /*
            foreach (object child in img.Children)
            {
                if (child is Button)
                {
                    if ((child as Button).Tag.ToString() == nameToLook)
                    {
                        Button foundButton = child as Button;
                        if (inForceAlignMode == false)
                        {
                            if (isLinkedPage == false)
                                foundButton.Visibility = Visibility.Visible;
                            else
                            {
                                nameToLook = "RemoveForceAlign" + splittedName[1];

                                foreach (object child2 in img.Children)
                                {
                                    if (child2 is Button)
                                    {
                                        if ((child2 as Button).Tag.ToString() == nameToLook)
                                        {
                                            foundButton = child2 as Button;
                                            
                                            if (isLinkedPage == true)
                                            {
                                                foundButton.ToolTip = "Page linked. Please remove existing link first.";
                                            }
                                            
                                            foundButton.Visibility = Visibility.Visible;
                                        }
                                        else
                                        {
                                            foundButton.Visibility = Visibility.Hidden;
                                        }
                                    }
                                }
                                
                            }
                        }
                        else
                        {
                            if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                            {
                                if (nameToLook.Contains("Left"))
                                {
                                    if (nameToLook != selectedSideGridButtonName1)
                                        foundButton.Visibility = Visibility.Hidden;
                                    else
                                        foundButton.Visibility = Visibility.Visible;
                                }
                                else
                                {
                                    if (lowerIndRight <= selfInd && selfInd <= upperIndLeft && isLinkedPage == false)
                                        foundButton.Visibility = Visibility.Visible;
                                    else
                                    {
                                        nameToLook = "SideButtonInvalidRight" + splittedName[1];
                                        foreach (object child2 in img.Children)
                                        {
                                            if (child2 is Button)
                                            {
                                                if ((child2 as Button).Tag.ToString() == nameToLook)
                                                {
                                                    foundButton = child2 as Button;

                                                    if (isLinkedPage == true)
                                                    {
                                                        foundButton.ToolTip = "Page linked. Please remove existing link first.";
                                                    }
                                                    else
                                                    {
                                                        foundButton.ToolTip = "This link would cross a previously set link. Please remove that link before aligning these pages.";
                                                    }

                                                    foundButton.Visibility = Visibility.Visible;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (nameToLook.Contains("Right"))
                                {
                                    if (nameToLook != selectedSideGridButtonName1)
                                        foundButton.Visibility = Visibility.Hidden;
                                    else
                                        foundButton.Visibility = Visibility.Visible;
                                }
                                else
                                {
                                    if (lowerIndRight <= selfInd && selfInd <= upperIndRight && isLinkedPage == false)
                                        foundButton.Visibility = Visibility.Visible;
                                    else
                                    {
                                        nameToLook = "SideButtonInvalidLeft" + splittedName[1];
                                        foreach (object child2 in img.Children)
                                        {
                                            if (child2 is Button)
                                            {
                                                if ((child2 as Button).Tag.ToString() == nameToLook)
                                                {
                                                    foundButton = child2 as Button;

                                                    if (isLinkedPage == true)
                                                    {
                                                        foundButton.ToolTip = "Page linked. Please remove existing link first.";
                                                    }
                                                    else
                                                    {
                                                        foundButton.ToolTip = "This link would cross a previously set link. Please remove that link before aligning these pages.";
                                                    }

                                                    foundButton.Visibility = Visibility.Visible;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            */

        }

        private void SideGridMouseLeave(object sender, MouseEventArgs args)
        {
            try
            {
                Grid img = sender as Grid;
                string imgName = img.Tag.ToString();
                string[] splittedName;
                string nameToLook = "";
                string nameToLook2 = "";
                string nameToLook3 = "";
                bool leftRight = false;

                if (imgName.Contains("Left"))
                {
                    splittedName = imgName.Split("Left");
                    nameToLook = "SideButtonLeft" + splittedName[1];
                    nameToLook2 = "SideButtonInvalidLeft" + splittedName[1];
                    nameToLook3 = "RemoveForceAlign" + splittedName[1];
                }
                else
                {
                    splittedName = imgName.Split("Right");
                    nameToLook = "SideButtonRight" + splittedName[1];
                    nameToLook2 = "SideButtonInvalidRight" + splittedName[1];
                    nameToLook3 = "RemoveForceAlign" + splittedName[1];
                }

                foreach (object child in img.Children)
                {
                    if (child is Button)
                    {
                        if ((child as Button).Tag.ToString() == nameToLook || (child as Button).Tag.ToString() == nameToLook2 || (child as Button).Tag.ToString() == nameToLook3)
                        {
                            Button foundButton = child as Button;
                            if (inForceAlignMode == false)
                            {
                                if ((child as Button).Tag.ToString() == nameToLook3)
                                {
                                    if (docs.forceAlignmentIndices.Count != 0)
                                    {
                                        if (leftRight == false)
                                        {
                                            foreach (List<int> l in docs.forceAlignmentIndices)
                                            {
                                                if (l[0] == docs.documents[docs.documentsToCompare[0]].docCompareIndices[int.Parse(splittedName[1])])
                                                {
                                                    foundButton.Visibility = Visibility.Visible;
                                                    break;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            foreach (List<int> l in docs.forceAlignmentIndices)
                                            {
                                                if (l[1] == docs.documents[docs.documentsToCompare[1]].docCompareIndices[int.Parse(splittedName[1])])
                                                {
                                                    foundButton.Visibility = Visibility.Visible;
                                                    break;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        foundButton.Visibility = Visibility.Hidden;
                                    }
                                }
                                else
                                {
                                    foundButton.Visibility = Visibility.Hidden;
                                }
                            }
                            else
                            {
                                if (foundButton.Tag.ToString() == selectedSideGridButtonName1)
                                {
                                    foundButton.Visibility = Visibility.Visible;
                                }
                                else if ((child as Button).Tag.ToString() == nameToLook3)
                                {
                                    if (docs.forceAlignmentIndices.Count != 0)
                                    {
                                        if (leftRight == false)
                                        {
                                            foreach (List<int> l in docs.forceAlignmentIndices)
                                            {
                                                if (l[0] == docs.documents[docs.documentsToCompare[0]].docCompareIndices[int.Parse(splittedName[1])])
                                                {
                                                    foundButton.Visibility = Visibility.Visible;
                                                    break;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            foreach (List<int> l in docs.forceAlignmentIndices)
                                            {
                                                if (l[1] == docs.documents[docs.documentsToCompare[1]].docCompareIndices[int.Parse(splittedName[1])])
                                                {
                                                    foundButton.Visibility = Visibility.Visible;
                                                    break;
                                                }
                                                else
                                                {
                                                    foundButton.Visibility = Visibility.Hidden;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        foundButton.Visibility = Visibility.Hidden;
                                    }
                                }
                                else
                                {
                                    foundButton.Visibility = Visibility.Hidden;
                                }
                            }
                        }
                    }
                }
            }
            catch
            {

            }
        }

        private void SidePanelDocCompareButton_Click(object sender, RoutedEventArgs e)
        {
            if (docs.documents.Count >= 2 && docCompareRunning == false && currentVisiblePanel != SidePanels.DOCCOMPARE)
            {
                if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPARETAB)
                {
                    PopupCompareDocBubble.IsOpen = false;
                    walkthroughStep = WalkthroughSteps.COMPAREHIGHLIGHT; // 7
                }

                inForceAlignMode = false;
                //DocCompareLeftStatsGrid.Visibility = Visibility.Collapsed;
                //DocCompareRightStatsGrid.Visibility = Visibility.Collapsed;

                SelectDocToCompare win = new SelectDocToCompare();
                List<string> filenames = new List<string>();
                for(int i = 0; i< docs.documents.Count; i++)
                {
                    if( i != docs.documentsToShow[0])
                    {
                        filenames.Add(Path.GetFileName(docs.documents[i].filePath));
                    }
                }

                if (docs.documents.Count > 2)
                {
                    win.Setup(filenames);
                    if (win.ShowDialog() == true)
                    {
                        docs.documentsToCompare[0] = docs.documentsToShow[0];

                        for (int i = 0; i < docs.documents.Count; i++)
                        {
                            if (Path.GetFileName(docs.documents[i].filePath) == filenames[win.selectedIndex])
                            {
                                docs.documentsToCompare[1] = i;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    docs.documentsToCompare[0] = docs.documentsToShow[0];
                    docs.documentsToCompare[1] = docs.documentsToShow[1];
                }

                //ReleaseDocPreview();
                UpdateDocCompareComboBox();
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                docs.forceAlignmentIndices = new List<List<int>>();
                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                //DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                //DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                //DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
                
            }
        }

        private void SidePanelOpenDocButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisiblePanel(SidePanels.DRAGDROP);

            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILETAB)
            {
                PopupBrowseFileBubble.IsOpen = false;
                PopupBrowseFileButtonBubble.IsOpen = true;
                walkthroughStep = WalkthroughSteps.BROWSEFILEBUTTON1;
            }
        }

        private void UnMaskSideGridFromForceAlignMode()
        {
            foreach (SideGridItemLeft item in DocCompareSideListViewLeft.Items)
            {
                item.GridEffect = null;
            }

            foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
            {
                item.GridEffect = null;
            }
        }

        private void UpdateDocCompareComboBox()
        {
            // update combo box left
            ObservableCollection<string> items = new ObservableCollection<string>();
            int selectedDocInd = 0;
            for (int i = 0; i < docs.documents.Count; i++)
            {
                if (i != docs.documentsToCompare[0] && docs.documents[i].processed == true)
                {
                    items.Add(Path.GetFileName(docs.documents[i].filePath));

                    if (i == docs.documentsToCompare[1])
                        selectedDocInd = items.Count - 1;
                }
            }

            DocCompareNameLabel2ComboBox.ItemsSource = items;
            DocCompareNameLabel2ComboBox.SelectedIndex = selectedDocInd;
        }

        private void UpdateDocSelectionComboBox()
        {
            // update combo box left
            ObservableCollection<string> items = new ObservableCollection<string>();
            int ind = 0;

            for (int i = 0; i < docs.documents.Count; i++)
            {
                bool ok = true;
                /*
                for (int j = 0; j < docs.documentsToShow.Count; j++)
                {
                    if (j != 0)
                    {
                        if (i != docs.documentsToShow[j])
                        {
                            ok &= true;
                        }
                        else
                        {
                            ok &= false;
                        }
                    }
                }
                */
                if (ok == true)
                    items.Add(Path.GetFileName(docs.documents[i].filePath));

                if (i == docs.documentsToShow[0])
                {
                    ind = items.Count - 1;
                }
            }
            Doc1NameLabelComboBox.ItemsSource = items;
            Doc1NameLabelComboBox.SelectedIndex = ind;

            // update combo box middle
            /*
            items = new ObservableCollection<string>();
            for (int i = 0; i < docs.documents.Count; i++)
            {
                bool ok = true;
                for (int j = 0; j < docs.documentsToShow.Count; j++)
                {
                    if (j != 1)
                    {
                        if (i != docs.documentsToShow[j])
                        {
                            ok &= true;
                        }
                        else
                        {
                            ok &= false;
                        }
                    }
                }

                if (ok == true)
                    items.Add(Path.GetFileName(docs.documents[i].filePath));

                if (i == docs.documentsToShow[1])
                {
                    ind = items.Count - 1;
                }
            }
            Doc2NameLabelComboBox.ItemsSource = items;
            Doc2NameLabelComboBox.SelectedIndex = ind;

            // TODO: Premium
            //if (settings.numPanelsDragDrop == 3)
            {
                // update combo box right
                items = new ObservableCollection<string>();
                for (int i = 0; i < docs.documents.Count; i++)
                {
                    bool ok = true;
                    for (int j = 0; j < docs.documentsToShow.Count; j++)
                    {
                        if (j != 2)
                        {
                            if (i != docs.documentsToShow[j])
                            {
                                ok &= true;
                            }
                            else
                            {
                                ok &= false;
                            }
                        }
                    }

                    if (ok == true)
                        items.Add(Path.GetFileName(docs.documents[i].filePath));

                    if (i == docs.documentsToShow[2])
                    {
                        ind = items.Count - 1;
                    }
                }
                Doc3NameLabelComboBox.ItemsSource = items;
                Doc3NameLabelComboBox.SelectedIndex = ind;

                items = new ObservableCollection<string>();
                for (int i = 0; i < docs.documents.Count; i++)
                {
                    bool ok = true;
                    for (int j = 0; j < docs.documentsToShow.Count; j++)
                    {
                        if (j != 2)
                        {
                            if (i != docs.documentsToShow[j])
                            {
                                ok &= true;
                            }
                            else
                            {
                                ok &= false;
                            }
                        }
                    }

                    if (ok == true)
                        items.Add(Path.GetFileName(docs.documents[i].filePath));

                    if (i == docs.documentsToShow[3])
                    {
                        ind = items.Count - 1;
                    }
                }
                Doc4NameLabelComboBox.ItemsSource = items;
                Doc4NameLabelComboBox.SelectedIndex = ind;

                items = new ObservableCollection<string>();
                for (int i = 0; i < docs.documents.Count; i++)
                {
                    bool ok = true;
                    for (int j = 0; j < docs.documentsToShow.Count; j++)
                    {
                        if (j != 2)
                        {
                            if (i != docs.documentsToShow[j])
                            {
                                ok &= true;
                            }
                            else
                            {
                                ok &= false;
                            }
                        }
                    }

                    if (ok == true)
                        items.Add(Path.GetFileName(docs.documents[i].filePath));

                    if (i == docs.documentsToShow[4])
                    {
                        ind = items.Count - 1;
                    }
                }
                Doc5NameLabelComboBox.ItemsSource = items;
                Doc5NameLabelComboBox.SelectedIndex = ind;
            }
            */
        }

        private void UpdateFileStat(int i)
        {
            if (i == 0 && docs.documents.Count >= 1) // DOC1
            {
                Doc1StatAuthorLabel.Text = docs.documents[docs.documentsToShow[0]].Creator;
                Doc1StatCreatedLabel.Text = docs.documents[docs.documentsToShow[0]].CreatedDate;
                Doc1StatLastEditorLabel.Text = docs.documents[docs.documentsToShow[0]].LastEditor;
                Doc1StatModifiedLabel.Text = docs.documents[docs.documentsToShow[0]].ModifiedDate;
            }

            if (i == 1 && docs.documents.Count >= 2) // DOC2
            {
                Doc2StatAuthorLabel.Text = docs.documents[docs.documentsToShow[1]].Creator;
                Doc2StatCreatedLabel.Text = docs.documents[docs.documentsToShow[1]].CreatedDate;
                Doc2StatLastEditorLabel.Text = docs.documents[docs.documentsToShow[1]].LastEditor;
                Doc2StatModifiedLabel.Text = docs.documents[docs.documentsToShow[1]].ModifiedDate;
            }

            if (i == 2 && docs.documents.Count >= 3) // DOC3
            {
                Doc3StatAuthorLabel.Text = docs.documents[docs.documentsToShow[2]].Creator;
                Doc3StatCreatedLabel.Text = docs.documents[docs.documentsToShow[2]].CreatedDate;
                Doc3StatLastEditorLabel.Text = docs.documents[docs.documentsToShow[2]].LastEditor;
                Doc3StatModifiedLabel.Text = docs.documents[docs.documentsToShow[2]].ModifiedDate;
            }

            if (i == 3 && docs.documents.Count >= 4) // DOC4
            {
                Doc4StatAuthorLabel.Text = docs.documents[docs.documentsToShow[3]].Creator;
                Doc4StatCreatedLabel.Text = docs.documents[docs.documentsToShow[3]].CreatedDate;
                Doc4StatLastEditorLabel.Text = docs.documents[docs.documentsToShow[3]].LastEditor;
                Doc4StatModifiedLabel.Text = docs.documents[docs.documentsToShow[3]].ModifiedDate;
            }

            if (i == 4 && docs.documents.Count >= 5) // DOC5
            {
                Doc5StatAuthorLabel.Text = docs.documents[docs.documentsToShow[4]].Creator;
                Doc5StatCreatedLabel.Text = docs.documents[docs.documentsToShow[4]].CreatedDate;
                Doc5StatLastEditorLabel.Text = docs.documents[docs.documentsToShow[4]].LastEditor;
                Doc5StatModifiedLabel.Text = docs.documents[docs.documentsToShow[4]].ModifiedDate;
            }

            if (i == 5) // DOC compare
            {
                DocCompareLeftStatAuthorLabel.Text = docs.documents[docs.documentsToCompare[0]].Creator;
                DocCompareLeftStatCreatedLabel.Text = docs.documents[docs.documentsToCompare[0]].CreatedDate;
                DocCompareLeftStatLastEditorLabel.Text = docs.documents[docs.documentsToCompare[0]].LastEditor;
                DocCompareLeftStatModifiedLabel.Text = docs.documents[docs.documentsToCompare[0]].ModifiedDate;
                DocCompareRightStatAuthorLabel.Text = docs.documents[docs.documentsToCompare[1]].Creator;
                DocCompareRightStatCreatedLabel.Text = docs.documents[docs.documentsToCompare[1]].CreatedDate;
                DocCompareRightStatLastEditorLabel.Text = docs.documents[docs.documentsToCompare[1]].LastEditor;
                DocCompareRightStatModifiedLabel.Text = docs.documents[docs.documentsToCompare[1]].ModifiedDate;
            }
        }

        private void UserEmailTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if ((sender as TextBox).Text.Length == 0)
            {
                ValidEmailTick.Visibility = Visibility.Hidden;
                InvalidEmailTick.Visibility = Visibility.Hidden;
            }
            else
            {
                if (Helper.IsValidEmail((sender as TextBox).Text))
                {
                    ValidEmailTick.Visibility = Visibility.Visible;
                    InvalidEmailTick.Visibility = Visibility.Hidden;
                    if (LicenseKeyTextBox.Text.Length == 19)
                    {
                        ActivateLicenseButton.IsEnabled = true;
                    }
                    else
                    {
                        ActivateLicenseButton.IsEnabled = false;
                    }
                }
                else
                {
                    ValidEmailTick.Visibility = Visibility.Hidden;
                    InvalidEmailTick.Visibility = Visibility.Visible;
                    ActivateLicenseButton.IsEnabled = false;
                }
            }
        }

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            if (PopupBrowseFileBubble.IsOpen == true)
            {
                PopupBrowseFileBubble.HorizontalOffset++;
                PopupBrowseFileBubble.HorizontalOffset--;
            }

            if (PopupBrowseFileButtonBubble.IsOpen == true)
            {
                PopupBrowseFileButtonBubble.HorizontalOffset++;
                PopupBrowseFileButtonBubble.HorizontalOffset--;
            }

            if (PopupDocPreviewNameComboboxBubble.IsOpen == true)
            {
                PopupDocPreviewNameComboboxBubble.HorizontalOffset++;
                PopupDocPreviewNameComboboxBubble.HorizontalOffset--;
            }

            if (PopupDocPreviewInfoButtonBubble.IsOpen == true)
            {
                PopupDocPreviewInfoButtonBubble.HorizontalOffset++;
                PopupDocPreviewInfoButtonBubble.HorizontalOffset--;
            }

            if (PopupDocPreviewInfoButton2Bubble.IsOpen == true)
            {
                PopupDocPreviewInfoButton2Bubble.HorizontalOffset++;
                PopupDocPreviewInfoButton2Bubble.HorizontalOffset--;
            }

            if (PopupCompareDocBubble.IsOpen == true)
            {
                PopupCompareDocBubble.HorizontalOffset++;
                PopupCompareDocBubble.HorizontalOffset--;
            }

            if (PopupHighlightOffBubble.IsOpen == true)
            {
                PopupHighlightOffBubble.HorizontalOffset++;
                PopupHighlightOffBubble.HorizontalOffset--;
            }

            if (PopupOpenOriBubble.IsOpen == true)
            {
                PopupOpenOriBubble.HorizontalOffset++;
                PopupOpenOriBubble.HorizontalOffset--;
            }

            if (PopupReloadBubble.IsOpen == true)
            {
                PopupReloadBubble.HorizontalOffset++;
                PopupReloadBubble.HorizontalOffset--;
            }

            if (PopupLinkPageBubble.IsOpen == true)
            {
                PopupLinkPageBubble.HorizontalOffset++;
                PopupLinkPageBubble.HorizontalOffset--;
            }

            if (PopupUnlinkBubble.IsOpen == true)
            {
                PopupUnlinkBubble.HorizontalOffset++;
                PopupUnlinkBubble.HorizontalOffset--;
            }

            if (PopupAnimateBubble.IsOpen == true)
            {
                PopupAnimateBubble.HorizontalOffset++;
                PopupAnimateBubble.HorizontalOffset--;
            }
        }

        private void Doc1NameLabelComboBox_DropDownOpened(object sender, EventArgs e)
        {
            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILECOMBOBOX)
            {
                PopupDocPreviewNameComboboxBubble.IsOpen = false;
            }
        }

        private void Doc1NameLabelComboBox_DropDownClosed(object sender, EventArgs e)
        {
            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.BROWSEFILECOMBOBOX)
            {
                PopupDocPreviewInfoButtonBubble.IsOpen = true;
                walkthroughStep = WalkthroughSteps.BROWSEFILEINFOOPEN;
            }
        }

        private void OpenDoc2OriginalButton_Click_1(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToCompare[1]].filePath + "\"";
            fileopener.Start();

            if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPAREOPENEXTERN)
            {
                PopupOpenOriBubble.IsOpen = false;
                PopupReloadBubble.IsOpen = true;
                walkthroughStep = WalkthroughSteps.COMPARERELOAD;
            }
        }

        private void OpenDoc1OriginalButton_Click_1(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToCompare[0]].filePath + "\"";
            fileopener.Start();
        }

        private void RestartWalkthroughButton_Click(object sender, RoutedEventArgs e)
        {
            settings.shownWalkthrough = false;
            SaveSettings();

            SetVisiblePanel(SidePanels.DRAGDROP);
            while (docs.documents.Count != 0)
            {
                docs.RemoveDocument(docs.documentsToShow[0], 0);
            }

            HideDragDropZone2();
            Doc1Grid.Visibility = Visibility.Hidden;
            Doc2Grid.Visibility = Visibility.Hidden;
            //Doc1PageNumberLabel.Content = "";
            //Doc2PageNumberLabel.Content = "";

            DocPreviewStatGrid.Visibility = Visibility.Collapsed;

            threadStartWalkthrough = new Thread(new ThreadStart(ShowWalkthroughStartMessage));
            threadStartWalkthrough.Start();
        }

        private void ExtendTrialTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                string currText = (sender as TextBox).Text;
                currText = currText.ToUpper();

                if (trialKeyLastInputString == null)
                    trialKeyLastInputString = currText;

                if (trialKeyLastInputString.Length >= currText.Length)
                {
                    if (currText.Length == 4 || currText.Length == 9 || currText.Length == 14)
                        currText = currText.Remove(currText.Length - 1);
                }
                else
                {
                    if (currText.Length == 4 || currText.Length == 9 || currText.Length == 14)
                        currText += '-';
                }

                if (currText.Length > 19)
                {
                    currText = currText.Remove(currText.Length - 1);
                }

                (sender as TextBox).Text = currText;
                (sender as TextBox).Select(currText.Length, 0);

                if (currText.Length == 19)
                {
                    ValidTrialKeyTick.Visibility = Visibility.Visible;
                    InvalidTrialKeyTick.Visibility = Visibility.Hidden;
                    ExtendTrialButton.IsEnabled = true;
                }
                else
                {
                    ValidTrialKeyTick.Visibility = Visibility.Hidden;
                    InvalidTrialKeyTick.Visibility = Visibility.Visible;
                    ExtendTrialButton.IsEnabled = false;
                }

                if (currText.Length == 0)
                {
                    ValidKeyTick.Visibility = Visibility.Hidden;
                    InvalidKeyTick.Visibility = Visibility.Hidden;
                    ActivateLicenseButton.IsEnabled = false;
                }

                trialKeyLastInputString = currText;
            }
            catch (Exception ex)
            {
                ErrorHandling.ReportException(ex);
            }
        }

        private void ExtendTrialButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ExtendTrialTextBox.Text == "QV9N-PQP4-5G9N-NV22")
                {
                    settings.showExtendTrial = false;
                    SaveSettings();
                    ExtendTrialGrid1.Visibility = Visibility.Hidden;
                    ExtendTrialGrid2.Visibility = Visibility.Hidden;
                    threadCheckTrial = new Thread(new ThreadStart(ExtendTrial));
                    threadCheckTrial.Start();
                    ValidTrialKeyTick.Visibility = Visibility.Hidden;
                    InvalidTrialKeyTick.Visibility = Visibility.Hidden;
                }
                else
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("Invalid code", "An invalid trial extension code entered.", "Okay");
                    msgBox.ShowDialog();
                }
            }
            catch
            {
            }
        }

        private void WindowUpdateButton_Click(object sender, RoutedEventArgs e)
        {
            CustomMessageBox msgBox = new CustomMessageBox();
            msgBox.Setup("Update avalaible", "A newer version of 2|Compare is available. Click OKAY to proceed with downloading the installer.", "Okay", "Skip");

            if (msgBox.ShowDialog() == false)
            {
                if (updateInstallerURL != null)
                {
                    ProcessStartInfo info = new ProcessStartInfo(updateInstallerURL)
                    {
                        UseShellExecute = true
                    };
                    Process.Start(info);
                }
            }
        }

        private void ChangeLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            CustomMessageBox msgBox = new CustomMessageBox();
            msgBox.Setup("Change License", "Proceed with activating another license? Current license will be deactivated.", "No", "Yes");
            if (msgBox.ShowDialog() == true)
            {
                Dispatcher.Invoke(() =>
                {
                    ActivateLicenseButton.IsEnabled = true;
                    ActivateLicenseButton.Visibility = Visibility.Visible;
                    ChangeLicenseButton.Visibility = Visibility.Hidden;

                    UserEmailTextBox.IsEnabled = true;
                    LicenseKeyTextBox.IsEnabled = true;
                });
            }
        }

        private void LinkScrollButton_Click(object sender, RoutedEventArgs e)
        {
            linkscroll = true;
            UnlinkScrollButton.Visibility = Visibility.Visible;
            LinkScrollButton.Visibility = Visibility.Hidden;

            // trigger a scroll
            Border border = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
            ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

            Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
            ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

            if (scrollViewer.ScrollableHeight > scrollViewer2.ScrollableHeight)
            {
                scrollViewer2.ScrollToVerticalOffset(scrollViewer2.VerticalOffset + 1);
                scrollViewer2.ScrollToVerticalOffset(scrollViewer2.VerticalOffset - 1);
            }
            else
            {
                scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset + 1);
                scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - 1);
            }
        }

        private void UnlinkScrollButton_Click(object sender, RoutedEventArgs e)
        {
            linkscroll = false;
            UnlinkScrollButton.Visibility = Visibility.Hidden;
            LinkScrollButton.Visibility = Visibility.Visible;
        }

        private void HandleDocPreviewMouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                Grid item = sender as Grid;
                Border child = item.Children[0] as Border;
                Image img = child.Child as Image;
                hiddenPPTEffect = img.Effect;
                img.Effect = null;

                hiddenPPTVisi = (item.Children[4] as Label).Visibility;
                System.Windows.Shapes.Path path = item.Children[3] as System.Windows.Shapes.Path;
                path.Visibility = Visibility.Hidden;
                Label label = item.Children[4] as Label;
                label.Visibility = Visibility.Hidden;
                Grid grid = item.Children[1] as Grid;
                grid.Visibility = Visibility.Hidden;

            }
            catch
            {

            }
        }

        private void HandleDocPreviewMouseLeave(object sender, MouseEventArgs e)
        {
            try
            {
                Grid item = sender as Grid;
                Border child = item.Children[0] as Border;
                Image img = child.Child as Image;
                img.Effect = hiddenPPTEffect;

                System.Windows.Shapes.Path path = item.Children[3] as System.Windows.Shapes.Path;
                path.Visibility = hiddenPPTVisi;
                Label label = item.Children[4] as Label;
                label.Visibility = hiddenPPTVisi;
                Grid grid = item.Children[1] as Grid;
                grid.Visibility = hiddenPPTVisi;

            }
            catch
            {

            }
        }

        private void HandleShowPPTNoteButton(object sender, RoutedEventArgs e)
        {
            try
            {
                Grid parentGrid = (sender as Button).Parent as Grid;
                foreach (object obj in parentGrid.Children)
                {
                    if (obj is Button && (obj as Button).Tag != null)
                    {
                        if ((obj as Button).Tag.ToString().Contains("ShowPPTNoteButton"))
                        {
                            (obj as Button).Visibility = Visibility.Hidden;
                        }
                    }

                    if (obj is Grid && (obj as Grid).Tag != null)
                    {
                        if ((obj as Grid).Tag.ToString().Contains("PPTSpeakerNoteGrid"))
                        {
                            (obj as Grid).Visibility = Visibility.Visible;
                        }
                    }
                }
            }
            catch
            {

            }
        }

        private void HandleClosePPTNoteButtonClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Grid containingGrid = (sender as Button).Parent as Grid;
                Grid parentGrid = containingGrid.Parent as Grid;
                foreach (object obj in parentGrid.Children)
                {
                    if (obj is Button && (obj as Button).Tag != null)
                    {
                        if ((obj as Button).Tag.ToString().Contains("ShowPPTNoteButton"))
                        {
                            (obj as Button).Visibility = Visibility.Visible;
                        }
                    }

                    if (obj is Grid && (obj as Grid).Tag != null)
                    {
                        if ((obj as Grid).Tag.ToString().Contains("PPTSpeakerNoteGrid"))
                        {
                            (obj as Grid).Visibility = Visibility.Hidden;
                        }
                    }
                }

            }
            catch
            {

            }
        }

        private void HandleClosePPTNoteButtonClickCompareMode(object sender, RoutedEventArgs e)
        {
            try
            {
                bool isChanged = false;
                string senderName = (sender as Button).Tag.ToString();
                string[] splitedName;

                if (senderName.Contains("Left"))
                    splitedName = senderName.Split("Left");
                else
                {
                    splitedName = senderName.Split("Right");
                }

                int i = int.Parse(splitedName[splitedName.Length - 1]);

                if (docs.pptSpeakerNotesDiff.Count != 0)
                {
                    if (docs.pptSpeakerNotesDiff[i].Count >= 1)
                        isChanged = true;

                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                    {
                        if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                        {
                            isChanged |= true;
                        }
                    }

                    if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                    {
                        if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                        {
                            isChanged |= true;
                        }
                    }
                }

                if (splitedName != null)
                {
                    CompareMainItem item = (CompareMainItem)DocCompareMainListView.Items[int.Parse(splitedName[splitedName.Length - 1])];

                    item.PPTNoteGridLeftVisi = Visibility.Hidden;
                    item.PPTNoteGridRightVisi = Visibility.Hidden;
                    item.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;

                    if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes != null)
                    {
                        if (isChanged == false)
                            item.showPPTSpeakerNotesButtonRight = Visibility.Visible;
                        else
                            item.showPPTSpeakerNotesButtonRightChanged = Visibility.Visible;
                    }
                    else
                    {
                        item.showPPTSpeakerNotesButtonLeft = Visibility.Visible;
                    }

                }
            }
            catch
            {

            }
        }

        private void HandleShowPPTNoteButtonCompareMode(object sender, RoutedEventArgs e)
        {
            try
            {
                bool isChanged = false;
                string senderName = (sender as Button).Tag.ToString();
                string[] splitedName;
                bool leftRight = false;

                if (senderName.Contains("Left"))
                    splitedName = senderName.Split("Left");
                else
                {

                    splitedName = senderName.Split("Right");
                    try
                    {
                        int i = int.Parse(splitedName[splitedName.Length - 1]);
                    }
                    catch
                    {
                        splitedName = senderName.Split("RightChanged");
                        isChanged = true;
                    }

                    leftRight = true;
                }

                if (splitedName != null)
                {
                    CompareMainItem item = (CompareMainItem)DocCompareMainListView.Items[int.Parse(splitedName[splitedName.Length - 1])];
                    if (leftRight == false)
                    {
                        item.PPTNoteGridLeftVisi = Visibility.Visible;
                        item.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                    }
                    else
                    {
                        if (item.PathToImgLeft != null)
                            item.PPTNoteGridLeftVisi = Visibility.Visible;
                        if (isChanged == false)
                            item.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                        else
                            item.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                        item.PPTNoteGridRightVisi = Visibility.Visible;
                    }
                }
            }
            catch
            {

            }
        }

        private void OpenDoc3OriginalButton_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToShow[2]].filePath + "\"";
            fileopener.Start();
        }

        private void ShowDoc4FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(4);

            if (Doc1StatAuthorLabel0.Visibility == Visibility.Collapsed)
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc1StatAuthorLabel.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Visible;
            }
            else
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc1StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Collapsed;
            }
        }

        private void ShowDoc5FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(5);

            if (Doc1StatAuthorLabel0.Visibility == Visibility.Collapsed)
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc1StatAuthorLabel.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc1StatCreatedLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc2StatAuthorLabel.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc2StatCreatedLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc4StatAuthorLabel.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc4StatCreatedLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc5StatAuthorLabel.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc5StatCreatedLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel.Visibility = Visibility.Visible;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Visible;
            }
            else
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc1StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc3StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc4StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc4StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc4StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc5StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc5StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc5StatLastEditorLabel0.Visibility = Visibility.Collapsed;
            }
        }

        private void DocCompareDragDropZone4_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void DocCompareDragDropZone4_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                string ext;

                if (data.Length > settings.maxDocCount)
                    ShowMaxDocCountWarningBox();

                foreach (string file in data)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    LoadFilesCommonPart();

                    //docs.documentsToShow[1] = docs.documents.Count - 1;

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }

        private void BrowseFileButton4_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif, *bmp)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF;*.bmp;*.BMP|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF,*.bmp,*.BMP |All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filenames = openFileDialog.FileNames;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);
                string ext;

                foreach (string file in filenames)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }
        /*
        private void Doc4NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc4NameLabelComboBox.SelectedItem.ToString();
                docs.documentsToShow[3] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayPreview(4, docs.documentsToShow[3]);
                UpdateDocSelectionComboBox();
                UpdateFileStat(3);
            }
            catch
            {
                Doc4NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
                UpdateFileStat(3);
            }
        }
        */

        private void OpenDoc4OriginalButton_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToShow[3]].filePath + "\"";
            fileopener.Start();
        }

        private void CloseDoc4Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[3], 3);
            CloseDocumentCommonPart();
            UpdateDocSelectionComboBox();

            if (docs.documents.Count < 3)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void DocCompareScrollViewer4_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                double accuHeight = 0;

                Border border = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                for (int i = 0; i < DocCompareListView4.Items.Count - 1; i++)
                {
                    ListViewItem container = DocCompareListView4.ItemContainerGenerator.ContainerFromItem(DocCompareListView4.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3 && Doc2Grid.Visibility == Visibility.Visible)
                    {
                        //Doc4PageNumberLabel.Content = (i + 1).ToString() + " / " + (DocCompareListView4.Items.Count-1).ToString();
                        break;
                    }
                    else
                    {
                        //Doc4PageNumberLabel.Content = "";
                    }
                }

                if (linkscroll == true)
                {
                    // try to scroll others
                    Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                    ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }
                }
            }
            catch
            {
            }
        }

        private void DocCompareDragDropZone5_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void DocCompareDragDropZone5_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                string ext;

                if (data.Length > settings.maxDocCount)
                    ShowMaxDocCountWarningBox();

                foreach (string file in data)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    LoadFilesCommonPart();

                    //docs.documentsToShow[1] = docs.documents.Count - 1;

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }

        private void BrowseFileButton5_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif, *bmp)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF;*.bmp;*.BMP|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF,*.bmp,*.BMP |All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filenames = openFileDialog.FileNames;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);
                string ext;

                foreach (string file in filenames)
                {
                    ext = Path.GetExtension(file);
                    if (ext != ".ppt" && ext != ".pptx" && ext != ".PPT" && ext != ".PPTX" && ext != ".pdf" && ext != ".PDF" && ext != ".jpg"
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG"
                        && ext != ".bmp" && ext != ".BMP")
                    {
                        ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(file));
                    }
                    else
                    {
                        if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                        {
                            docs.AddDocument(file);
                        }
                        else if (docs.documents.Count >= settings.maxDocCount)
                        {
                            ShowMaxDocCountWarningBox();
                            break;
                        }
                        else
                        {
                            ShowExistingDocCountWarningBox(file);
                        }
                    }
                }

                if (docs.documents.Count != 0)
                {
                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }
        /*
        private void Doc5NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc4NameLabelComboBox.SelectedItem.ToString();
                docs.documentsToShow[4] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayPreview(5, docs.documentsToShow[4]);
                UpdateDocSelectionComboBox();
                UpdateFileStat(4);
            }
            catch
            {
                Doc4NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
                UpdateFileStat(4);
            }
        }
        */
        private void OpenDoc5OriginalButton_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToShow[4]].filePath + "\"";
            fileopener.Start();
        }

        private void ReloadDoc4Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDoc4.Visibility = Visibility.Visible;

            docs.docToReload = docs.documentsToShow[3];
            docs.displayToReload = 3;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void CloseDoc5Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[4], 4);
            CloseDocumentCommonPart();
            UpdateDocSelectionComboBox();

            if (docs.documents.Count < 2)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void ReloadDoc5Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDoc5.Visibility = Visibility.Visible;

            docs.docToReload = docs.documentsToShow[4];
            docs.displayToReload = 4;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void DocCompareScrollViewer5_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                double accuHeight = 0;

                Border border = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

                for (int i = 0; i < DocCompareListView5.Items.Count - 1; i++)
                {
                    ListViewItem container = DocCompareListView5.ItemContainerGenerator.ContainerFromItem(DocCompareListView5.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3 && Doc2Grid.Visibility == Visibility.Visible)
                    {
                        //Doc5PageNumberLabel.Content = (i + 1).ToString() + " / " + (DocCompareListView5.Items.Count-1).ToString();
                        break;
                    }
                    else
                    {
                        //Doc5PageNumberLabel.Content = "";
                    }
                }

                if (linkscroll == true)
                {
                    // try to scroll others
                    Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                    ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }

                    border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                    scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                    }
                    else
                    {
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                    }
                }
            }
            catch
            {
            }
        }

        private void CompareDoc2Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                docs.documentsToCompare[0] = docs.documentsToShow[0]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[0]].processed == false)
                {
                    docs.documentsToCompare[0] = FindNextDocToShow();
                }

                docs.documentsToCompare[1] = docs.documentsToShow[1]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[1]].processed == false)
                {
                    docs.documentsToCompare[1] = FindNextDocToShow();
                }

                //ReleaseDocPreview();
                UpdateDocCompareComboBox();
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                docs.forceAlignmentIndices = new List<List<int>>();
                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                //DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                //DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                //DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
            catch
            {

            }
        }

        private void CompareDoc3Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                docs.documentsToCompare[0] = docs.documentsToShow[0]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[0]].processed == false)
                {
                    docs.documentsToCompare[0] = FindNextDocToShow();
                }

                docs.documentsToCompare[1] = docs.documentsToShow[2]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[1]].processed == false)
                {
                    docs.documentsToCompare[1] = FindNextDocToShow();
                }

                //ReleaseDocPreview();
                UpdateDocCompareComboBox();
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                docs.forceAlignmentIndices = new List<List<int>>();
                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                //DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                //DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                //DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
            catch
            {

            }
        }

        private void CompareDoc4Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                docs.documentsToCompare[0] = docs.documentsToShow[0]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[0]].processed == false)
                {
                    docs.documentsToCompare[0] = FindNextDocToShow();
                }

                docs.documentsToCompare[1] = docs.documentsToShow[3]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[1]].processed == false)
                {
                    docs.documentsToCompare[1] = FindNextDocToShow();
                }

                //ReleaseDocPreview();
                UpdateDocCompareComboBox();
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                docs.forceAlignmentIndices = new List<List<int>>();
                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                //DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                //DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                //DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
            catch
            {

            }
        }

        private void CompareDoc5Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                docs.documentsToCompare[0] = docs.documentsToShow[0]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[0]].processed == false)
                {
                    docs.documentsToCompare[0] = FindNextDocToShow();
                }

                docs.documentsToCompare[1] = docs.documentsToShow[4]; // default using the first document selected
                if (docs.documents[docs.documentsToCompare[1]].processed == false)
                {
                    docs.documentsToCompare[1] = FindNextDocToShow();
                }

                //ReleaseDocPreview();
                UpdateDocCompareComboBox();
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                docs.forceAlignmentIndices = new List<List<int>>();
                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                //DocCompareMainListView.ScrollIntoView(DocCompareMainListView.Items[0]);
                //DocCompareSideListViewLeft.ScrollIntoView(DocCompareSideListViewLeft.Items[0]);
                //DocCompareSideListViewRight.ScrollIntoView(DocCompareSideListViewRight.Items[0]);
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
            catch
            {

            }
        }

        private void Doc2Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            CompareDoc2Button.Visibility = Visibility.Visible;
        }

        private void Doc2Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc2Button.Visibility = Visibility.Hidden;
        }

        private void Doc3Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            CompareDoc3Button.Visibility = Visibility.Visible;
        }

        private void Doc3Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc3Button.Visibility = Visibility.Hidden;
        }

        private void Doc4Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            CompareDoc4Button.Visibility = Visibility.Visible;
        }

        private void Doc4Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc4Button.Visibility = Visibility.Hidden;
        }

        private void Doc5Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            CompareDoc5Button.Visibility = Visibility.Visible;
        }

        private void Doc5Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc5Button.Visibility = Visibility.Hidden;
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (PopupBrowseFileBubble.IsOpen == true)
            {
                PopupBrowseFileBubble.HorizontalOffset++;
                PopupBrowseFileBubble.HorizontalOffset--;
            }

            if (PopupBrowseFileButtonBubble.IsOpen == true)
            {
                PopupBrowseFileButtonBubble.HorizontalOffset++;
                PopupBrowseFileButtonBubble.HorizontalOffset--;
            }

            if (PopupDocPreviewNameComboboxBubble.IsOpen == true)
            {
                PopupDocPreviewNameComboboxBubble.HorizontalOffset++;
                PopupDocPreviewNameComboboxBubble.HorizontalOffset--;
            }

            if (PopupDocPreviewInfoButtonBubble.IsOpen == true)
            {
                PopupDocPreviewInfoButtonBubble.HorizontalOffset++;
                PopupDocPreviewInfoButtonBubble.HorizontalOffset--;
            }

            if (PopupDocPreviewInfoButton2Bubble.IsOpen == true)
            {
                PopupDocPreviewInfoButton2Bubble.HorizontalOffset++;
                PopupDocPreviewInfoButton2Bubble.HorizontalOffset--;
            }

            if (PopupCompareDocBubble.IsOpen == true)
            {
                PopupCompareDocBubble.HorizontalOffset++;
                PopupCompareDocBubble.HorizontalOffset--;
            }

            if (PopupHighlightOffBubble.IsOpen == true)
            {
                PopupHighlightOffBubble.HorizontalOffset++;
                PopupHighlightOffBubble.HorizontalOffset--;
            }

            if (PopupOpenOriBubble.IsOpen == true)
            {
                PopupOpenOriBubble.HorizontalOffset++;
                PopupOpenOriBubble.HorizontalOffset--;
            }

            if (PopupReloadBubble.IsOpen == true)
            {
                PopupReloadBubble.HorizontalOffset++;
                PopupReloadBubble.HorizontalOffset--;
            }

            if (PopupLinkPageBubble.IsOpen == true)
            {
                PopupLinkPageBubble.HorizontalOffset++;
                PopupLinkPageBubble.HorizontalOffset--;
            }

            if (PopupUnlinkBubble.IsOpen == true)
            {
                PopupUnlinkBubble.HorizontalOffset++;
                PopupUnlinkBubble.HorizontalOffset--;
            }

            if (PopupAnimateBubble.IsOpen == true)
            {
                PopupAnimateBubble.HorizontalOffset++;
                PopupAnimateBubble.HorizontalOffset--;
            }

            if (WindowState == WindowState.Normal)
            {
                outerBorder.Margin = new Thickness(0);
            }

            /*
            if(Doc1StatsGrid.ActualHeight > Doc2StatsGrid.ActualHeight)
                Doc2StatsGrid.Height = Doc1StatsGrid.ActualHeight;
            else
                Doc1StatsGrid.Height = Doc2StatsGrid.ActualHeight;
            */
        }

        /*
        private void ReleaseDocPreview()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        */

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Normal)
            {
                WindowMaximizeButton.Visibility = Visibility.Visible;
                WindowRestoreButton.Visibility = Visibility.Hidden;

            }

            if (WindowState == WindowState.Maximized)
            {
                WindowMaximizeButton.Visibility = Visibility.Hidden;
                WindowRestoreButton.Visibility = Visibility.Visible;
                outerBorder.Margin = new Thickness(5, 5, 5, 0);
            }
        }

        private void WindowCloseButton_Click(object sender, RoutedEventArgs e)
        {
            //TODO: implement handling for query before closing

            DirectoryInfo di;

            foreach (Document doc in docs.documents)
            {
                if (Directory.Exists(doc.imageFolder))
                {
                    di = new DirectoryInfo(doc.imageFolder);
                    di.Delete(true);
                }
            }

            if (Directory.Exists(compareResultFolder))
            {
                di = new DirectoryInfo(compareResultFolder);
                di.Delete(true);
            }

            di = new DirectoryInfo(appDataDir);
            foreach (DirectoryInfo di2 in di.EnumerateDirectories())
            {
                if (di2.Name != "lic")
                {
                    di2.Delete(true);
                }
            }

            // stop any thread still running
            
            Close();
        }

        private void WindowMaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowMaximizeButton.Visibility = Visibility.Hidden;
            WindowRestoreButton.Visibility = Visibility.Visible;
            WindowState = WindowState.Maximized;
            MaxHeight = SystemParameters.MaximizedPrimaryScreenHeight - 7;
            outerBorder.Margin = new Thickness(5, 5, 5, 0);
        }

        private void WindowMinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
            outerBorder.Margin = new Thickness(0);
        }

        private void WindowRestoreButton_Click(object sender, RoutedEventArgs e)
        {
            WindowMaximizeButton.Visibility = Visibility.Visible;
            WindowRestoreButton.Visibility = Visibility.Hidden;
            WindowState = WindowState.Normal;
            outerBorder.Margin = new Thickness(0);
        }
    }

    public class SideGridItemLeft : INotifyPropertyChanged
    {
        private Color _color;
        private Effect _effect;

        public event PropertyChangedEventHandler PropertyChanged;

        private Visibility _forceAlignButtonVisi;
        private Visibility _forceAlignButtonInvalidVisi;

        private Visibility _showHidden;

        public Visibility ShowHidden
        {
            get
            {
                return _showHidden;
            }

            set
            {
                _showHidden = value;
                OnPropertyChanged();
            }
        }
        public Visibility ForceAlignButtonVisi
        {
            get
            {
                return _forceAlignButtonVisi;
            }

            set
            {
                _forceAlignButtonVisi = value;
                OnPropertyChanged();
            }
        }
        public Visibility ForceAlignButtonInvalidVisi
        {
            get
            {
                return _forceAlignButtonInvalidVisi;
            }

            set
            {
                _forceAlignButtonInvalidVisi = value;
                OnPropertyChanged();
            }
        }

        public Color BackgroundBrush
        {
            get
            {
                return _color;
            }

            set
            {
                _color = value;
                OnPropertyChanged();
            }
        }

        public string ForceAlignButtonName { get; set; }
        public string ForceAlignInvalidButtonName { get; set; }

        public Effect GridEffect
        {
            get
            {
                return _effect;
            }

            set
            {
                _effect = value;
                OnPropertyChanged();
            }
        }

        public string GridName { get; set; }
        public string ImgDummyName { get; set; }
        public string ImgGridName { get; set; }
        public string ImgName { get; set; }
        public Thickness Margin { get; set; }
        public string PageNumberLabel { get; set; }

        private string _pathToImg;
        public string PathToImg { get { return _pathToImg; } set { _pathToImg = value; OnPropertyChanged(); } }

        private string _pathToImgDummy;
        public string PathToImgDummy { get { return _pathToImgDummy; } set { _pathToImgDummy = value; OnPropertyChanged(); } }
        public string RemoveForceAlignButtonName { get; set; }
        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SideGridItemRight : INotifyPropertyChanged
    {
        private Color _color;
        private Effect _effect;
        private Visibility _showMask;

        public event PropertyChangedEventHandler PropertyChanged;

        private Visibility _showHidden;
        private Visibility _forceAlignButtonVisi;
        private Visibility _forceAlignButtonInvalidVisi;

        private Visibility _diffVisi;
        private Visibility _noDiffVisi;

        public Visibility ShowHidden
        {
            get
            {
                return _showHidden;
            }

            set
            {
                _showHidden = value;
                OnPropertyChanged();
            }
        }

        public Visibility ForceAlignButtonVisi
        {
            get
            {
                return _forceAlignButtonVisi;
            }

            set
            {
                _forceAlignButtonVisi = value;
                OnPropertyChanged();
            }
        }
        public Visibility ForceAlignButtonInvalidVisi
        {
            get
            {
                return _forceAlignButtonInvalidVisi;
            }

            set
            {
                _forceAlignButtonInvalidVisi = value;
                OnPropertyChanged();
            }

        }
        public Visibility DiffVisi
        {
            get
            {
                return _diffVisi;
            }

            set
            {
                _diffVisi = value;
                OnPropertyChanged();
            }
        }

        public Visibility NoDiffVisi
        {
            get
            {
                return _noDiffVisi;
            }

            set
            {
                _noDiffVisi = value;
                OnPropertyChanged();
            }
        }
        public Color BackgroundBrush
        {
            get
            {
                return _color;
            }

            set
            {
                _color = value;
                OnPropertyChanged();
            }
        }

        public string ForceAlignButtonName { get; set; }
        public string ForceAlignInvalidButtonName { get; set; }

        public Effect GridEffect
        {
            get
            {
                return _effect;
            }

            set
            {
                _effect = value;
                OnPropertyChanged();
            }
        }

        public string GridName { get; set; }
        public string ImgDummyName { get; set; }
        public string ImgGridName { get; set; }
        public string ImgMaskName { get; set; }
        public string ImgName { get; set; }
        public Thickness Margin { get; set; }

        private string _pathToImg;
        public string PathToImg { get { return _pathToImg; } set { _pathToImg = value; OnPropertyChanged(); } }

        private string _pathToImgDummy;
        public string PathToImgDummy { get { return _pathToImgDummy; } set { _pathToImgDummy = value; OnPropertyChanged(); } }

        private string _pathToMask;
        public string PathToMask { get { return _pathToMask; } set { _pathToMask = value; OnPropertyChanged(); } }
        public bool RemoveForceAlignButtonEnable { get; set; }
        public string RemoveForceAlignButtonName { get; set; }

        public Visibility RemoveForceAlignButtonVisibility { get; set; }

        public Visibility ShowMask
        {
            get
            {
                return _showMask;
            }

            set
            {
                _showMask = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SimpleImageItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public Thickness Margin { get; set; }

        private string _pathToFile;
        private Visibility _eodVisi;
        private double _blurRadius;
        private Visibility _showHidden;
        private Visibility _showPPTNoteButton;

        public string ShowPPTSpeakerNotesButtonName { get; set; }
        public string HiddenPPTGridName { get; set; }
        public string PPTSpeakerNoteGridName { get; set; }
        public string ClosePPTSpeakerNotesButtonName { get; set; }
        public string PPTSpeakerNotes { get; set; }


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

        public Visibility EoDVisi
        {
            get
            {
                return _eodVisi;
            }

            set
            {
                _eodVisi = value;
                OnPropertyChanged();
            }
        }

        public Visibility showHidden
        {
            get
            {
                return _showHidden;
            }

            set
            {
                _showHidden = value;
                OnPropertyChanged();
            }
        }
        public Visibility showPPTSpeakerNotesButton
        {
            get
            {
                return _showPPTNoteButton;
            }

            set
            {
                _showPPTNoteButton = value;
                OnPropertyChanged();
            }
        }

        public double BlurRadius
        {
            get
            {
                return _blurRadius;
            }

            set
            {
                _blurRadius = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class UriToCachedImageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value == null)
                return null;

            if (!string.IsNullOrEmpty(value.ToString()))
            {
                BitmapImage bi = new BitmapImage();
                bi.BeginInit();
                bi.UriSource = new Uri(value.ToString());
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.EndInit();
                return bi;
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException("Two way conversion is not supported.");
        }
    }
}