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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;

namespace DocCompareWPF
{
    public class CompareMainItem : INotifyPropertyChanged
    {
        private Visibility _showMask;

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

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string appDataDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare");
        private readonly DocumentManagement docs;
        private readonly string versionString = "1.0.6";
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
        private WalkthroughSteps walkthroughStep = 0;

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
                        if (child is Image)
                        {
                            Image thisImg = child as Image;
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
                    if ((sender as Button).Name.Contains("Top"))
                    {
                        if (docs.documents.Count >= 2)
                        {
                            docs.documentsToShow[1] = docs.documents.Count - 1;
                        }
                    }

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
            if (docs.documentsToShow[0] != -1)
            {
                DisplayImageLeft(docs.documentsToShow[0]);

                if (docs.documents.Count >= 2)
                {
                    if (docs.documentsToShow[1] != -1)
                        DisplayImageMiddle(docs.documentsToShow[1]);

                    if (docs.documentsToShow[1] == -1)
                    {
                        Doc2Grid.Visibility = Visibility.Hidden;
                        DocCompareDragDropZone2.Visibility = Visibility.Visible;
                        DocPreviewStatGrid.Visibility = Visibility.Visible;
                        Doc1StatsGrid.Visibility = Visibility.Visible;
                        Doc2StatsGrid.Visibility = Visibility.Collapsed;
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
                        Doc1PageNumberLabel.Content = "";
                        //Doc2StatsGrid.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    Doc2Grid.Visibility = Visibility.Hidden;
                    Doc2PageNumberLabel.Content = "";
                    DocCompareDragDropZone2.Visibility = Visibility.Visible;
                    DocPreviewStatGrid.Visibility = Visibility.Visible;
                    Doc1StatsGrid.Visibility = Visibility.Visible;
                    Doc2StatsGrid.Visibility = Visibility.Collapsed;
                    ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                    ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
                    Doc2PageNumberLabel.Content = "";
                    //Doc2StatsGrid.Visibility = Visibility.Collapsed;
                }

                if (docs.documents.Count >= 3 && settings.numPanelsDragDrop == 3)
                {
                    if (docs.documentsToShow[2] != -1)
                        DisplayImageRight(docs.documentsToShow[2]);

                    if (docs.documentsToShow[2] == -1)
                    {
                        HideDragDropZone3();
                        //Doc3StatsGrid.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    if (settings.numPanelsDragDrop == 3 && docs.documents.Count >= 2)
                    {
                        Doc3Grid.Visibility = Visibility.Hidden;
                        DocCompareDragDropZone3.Visibility = Visibility.Visible;
                        ShowDragDropZone3();
                    }
                    else
                    {
                        HideDragDropZone3();
                        //Doc3StatsGrid.Visibility = Visibility.Collapsed;
                    }
                }
            }
            else
            {
                Doc1Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone1.Visibility = Visibility.Visible;
                DocPreviewStatGrid.Visibility = Visibility.Collapsed;
                Doc1PageNumberLabel.Content = "";
                Doc2PageNumberLabel.Content = "";
                Doc1StatsGrid.Visibility = Visibility.Visible;
                Doc2StatsGrid.Visibility = Visibility.Visible;
                ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
            }

            if (docs.documents.Count == 0)
            {
                Doc1Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone1.Visibility = Visibility.Visible;
                DocPreviewStatGrid.Visibility = Visibility.Collapsed;
                Doc1StatsGrid.Visibility = Visibility.Visible;
                Doc2StatsGrid.Visibility = Visibility.Visible;
                ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
                Doc1PageNumberLabel.Content = "";
                Doc2PageNumberLabel.Content = "";
                HideDragDropZone2();
                HideDragDropZone3();
            }
        }

        private void CompareDocsThread()
        {
            try
            {
                Dispatcher.Invoke(() => { UpdateFileStat(3); });

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
                };

                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                {
                    thisItem.PathToImgLeft = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg");

                    if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                    {
                        thisItem.PathToAniImgLeft = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg");
                        thisItem.AniDiffButtonEnable = true;
                    }
                }

                if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                {
                    thisItem.PathToImgRight = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg");

                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                    {
                        thisItem.PathToAniImgRight = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg");

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
                        rightItem.RemoveForceAlignButtonName = "RemoveForceAlign" + docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString();
                        rightItem.RemoveForceAlignButtonVisibility = Visibility.Visible;
                        rightItem.RemoveForceAlignButtonEnable = true;
                    }
                }

                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                {
                    leftItem.PathToImg = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg");
                    rightItem.PathToImgDummy = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg");
                }

                if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1) // doc 2 has a valid page
                {
                    rightItem.PathToImg = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg");
                    leftItem.PathToImgDummy = Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg");

                    if (File.Exists(Path.Join(compareResultFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png")))
                    {
                        rightItem.PathToMask = Path.Join(compareResultFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");
                        if (i != 0)
                        {
                            leftItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            rightItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                        }

                        if (showMask == true)
                            rightItem.ShowMask = Visibility.Visible;
                        else
                            rightItem.ShowMask = Visibility.Hidden;
                    }
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
                                    PathToFile = Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".jpg")
                                };

                                if (pageCounter == 0)
                                {
                                    thisImage.Margin = new Thickness(10, 10, 10, 10);
                                }
                                else
                                {
                                    thisImage.Margin = new Thickness(10, 0, 10, 10);
                                }
                                imageList.Add(thisImage);
                                pageCounter++;
                            }
                        }
                        else
                        {
                            /*
                            Grid errGrid = new Grid()
                            {
                                HorizontalAlignment = HorizontalAlignment.Stretch,
                                VerticalAlignment = VerticalAlignment.Stretch,
                                Height = DocCompareScrollViewer1.ActualHeight,
                            };

                            Tile errCard = new Tile()
                            {
                                HorizontalAlignment = HorizontalAlignment.Center,
                                VerticalAlignment = VerticalAlignment.Center,
                                Padding = new Thickness(10),
                                Width = 250,
                                Background = FindResource("SecondaryAccentBrush") as Brush,
                            };

                            Label errLabel = new Label()
                            {
                                Foreground = Brushes.White,
                                Content = "There was an error loading this file",
                            };

                            errCard.Content = errLabel;

                            errGrid.Children.Add(errCard);
                            childPanel1.Children.Add(errGrid);
                            childPanel1.Height = DocCompareScrollViewer1.ActualHeight;
                            */
                        }

                        DocCompareListView1.ItemsSource = imageList;
                        DocCompareListView1.Items.Refresh();
                        DocCompareListView1.ScrollIntoView(DocCompareListView1.Items[0]);
                        Doc1Grid.Visibility = Visibility.Visible;
                        ProgressBarDoc1.Visibility = Visibility.Hidden;
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
                                        PathToFile = Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".jpg")
                                    };

                                    if (pageCounter == 0)
                                    {
                                        thisImage.Margin = new Thickness(10, 10, 10, 10);
                                    }
                                    else
                                    {
                                        thisImage.Margin = new Thickness(10, 0, 10, 10);
                                    }
                                    imageList.Add(thisImage);
                                    pageCounter++;
                                }
                            }
                            else
                            {
                                /*
                                Grid errGrid = new Grid()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Stretch,
                                    VerticalAlignment = VerticalAlignment.Stretch,
                                    Height = DocCompareScrollViewer2.ActualHeight,
                                };

                                Tile errCard = new Tile()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Center,
                                    VerticalAlignment = VerticalAlignment.Center,
                                    Padding = new Thickness(10),
                                    Width = 250,
                                    Background = FindResource("SecondaryAccentBrush") as Brush,
                                };

                                Label errLabel = new Label()
                                {
                                    Foreground = Brushes.White,
                                    Content = "There was an error loading this file",
                                };

                                errCard.Content = errLabel;

                                errGrid.Children.Add(errCard);
                                childPanel2.Children.Add(errGrid);
                                childPanel2.Height = DocCompareScrollViewer2.ActualHeight;
                                */
                            }

                            DocCompareListView2.ItemsSource = imageList;
                            DocCompareListView2.Items.Refresh();
                            DocCompareListView2.ScrollIntoView(DocCompareListView2.Items[0]);
                            Doc2Grid.Visibility = Visibility.Visible;
                            ProgressBarDoc2.Visibility = Visibility.Hidden;
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
                                        PathToFile = Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".jpg")
                                    };

                                    if (pageCounter == 0)
                                    {
                                        thisImage.Margin = new Thickness(10, 10, 10, 10);
                                    }
                                    else
                                    {
                                        thisImage.Margin = new Thickness(10, 0, 10, 10);
                                    }
                                    imageList.Add(thisImage);
                                    pageCounter++;
                                }
                            }
                            else
                            {
                                /*
                                Grid errGrid = new Grid()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Stretch,
                                    VerticalAlignment = VerticalAlignment.Stretch,
                                    Height = DocCompareScrollViewer3.ActualHeight,
                                };

                                Tile errCard = new Tile()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Center,
                                    VerticalAlignment = VerticalAlignment.Center,
                                    Padding = new Thickness(10),
                                    Width = 250,
                                    Background = FindResource("SecondaryAccentBrush") as Brush,
                                };

                                Label errLabel = new Label()
                                {
                                    Foreground = Brushes.White,
                                    Content = "There was an error loading this file",
                                };

                                errCard.Content = errLabel;

                                errGrid.Children.Add(errCard);
                                childPanel3.Children.Add(errGrid);
                                childPanel3.Height = DocCompareScrollViewer3.ActualHeight;
                                */
                            }

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
                        var stream = File.OpenRead(Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".jpg"));
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
                docs.documentsToShow[0] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayImageLeft(docs.documentsToShow[0]);
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

        private void Doc2NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc2NameLabelComboBox.SelectedItem.ToString();
                docs.documentsToShow[1] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayImageMiddle(docs.documentsToShow[1]);
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
                DisplayImageRight(docs.documentsToShow[2]);
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
                    docs.documentsToShow[0] = docs.documents.Count - 1;

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

                    docs.documentsToShow[1] = docs.documents.Count - 1;

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

                    docs.documentsToShow[2] = docs.documents.Count - 1;

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

                UpdateFileStat(3);
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

                for (int i = 0; i < DocCompareListView1.Items.Count; i++)
                {
                    ListViewItem container = DocCompareListView1.ItemContainerGenerator.ContainerFromItem(DocCompareListView1.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3)
                    {
                        Doc1PageNumberLabel.Content = (i + 1).ToString() + " / " + DocCompareListView1.Items.Count.ToString();
                        break;
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

                for (int i = 0; i < DocCompareListView2.Items.Count; i++)
                {
                    ListViewItem container = DocCompareListView2.ItemContainerGenerator.ContainerFromItem(DocCompareListView2.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3)
                    {
                        Doc2PageNumberLabel.Content = (i + 1).ToString() + " / " + DocCompareListView2.Items.Count.ToString();
                        break;
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

                for (int i = 0; i < DocCompareListView3.Items.Count; i++)
                {
                    ListViewItem container = DocCompareListView3.ItemContainerGenerator.ContainerFromItem(DocCompareListView3.Items[i]) as ListViewItem;

                    accuHeight += container.ActualHeight;

                    if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3)
                    {
                        Doc3PageNumberLabel.Content = (i + 1).ToString() + " / " + DocCompareListView3.Items.Count.ToString();
                        break;
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
                    item.ShowMask = Visibility.Hidden;
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
                    if (child is Image)
                    {
                        Image thisImg = child as Image;
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

            foreach (CompareMainItem item in DocCompareMainListView.Items)
            {
                if (showMask == true)
                {
                    item.ShowMask = Visibility.Visible;
                }
                else
                {
                    item.ShowMask = Visibility.Hidden;
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
                            if (thisButton.Tag.ToString().Contains("Left"))
                                thisButton.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            if (thisButton.Tag.ToString().Contains("Right"))
                                thisButton.Visibility = Visibility.Visible;
                        }
                    }
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
                            if (thisButton.Tag.ToString().Contains("Left"))
                                thisButton.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            if (thisButton.Tag.ToString().Contains("Right"))
                                thisButton.Visibility = Visibility.Hidden;
                        }
                    }
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
        }

        private void HideDragDropZone3()
        {
            DocCompareThirdDocZone.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[2].Width = new GridLength(0, GridUnitType.Star);
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
                Application.Current.Dispatcher.Invoke(() =>
                {
                    Doc1Grid.Visibility = Visibility.Hidden;
                    ProgressBarDoc1.Visibility = Visibility.Visible;
                    ShowDragDropZone2();
                });
            }

            if (docs.documents.Count >= 2)
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

            if (docs.documents.Count >= 3)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    Doc3Grid.Visibility = Visibility.Hidden;
                    ProgressBarDoc3.Visibility = Visibility.Visible;
                });
            }

            docProcessRunning = true;

            Application.Current.Dispatcher.Invoke(() =>
            {
                ProcessingDocProgressCard.Visibility = Visibility.Visible;
                ProcessingDocProgressbar.Value = 0;
                ProcessingDocLabel.Text = "Processing: " + Path.GetFileName(docs.documents[0].filePath);

                BrowseFileTopButton1.IsEnabled = false;
                BrowseFileTopButton2.IsEnabled = false;
                BrowseFileTopButton3.IsEnabled = false;

                ReloadDoc1Button.IsEnabled = false;
                ReloadDoc2Button.IsEnabled = false;
                ReloadDoc3Button.IsEnabled = false;

                CloseDoc1Button.IsEnabled = false;
                CloseDoc2Button.IsEnabled = false;
                CloseDoc3Button.IsEnabled = false;

                OpenDoc1OriginalButton1.IsEnabled = false;
                OpenDoc2OriginalButton2.IsEnabled = false;
                OpenDoc3OriginalButton3.IsEnabled = false;

                ShowDoc1FileInfoButton.IsEnabled = false;
                ShowDoc2FileInfoButton.IsEnabled = false;
                ShowDoc3FileInfoButton.IsEnabled = false;
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

        private void OpenDoc3OriginalButton3_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToShow[2]].filePath + "\"";
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
                        BrowseFileTopButton1.IsEnabled = false;
                        BrowseFileTopButton2.IsEnabled = false;
                        BrowseFileTopButton3.IsEnabled = false;
                        BrowseFileButton1.IsEnabled = false;
                        BrowseFileButton2.IsEnabled = false;
                        BrowseFileButton3.IsEnabled = false;

                        DocCompareFirstDocZone.AllowDrop = false;
                        DocCompareDragDropZone1.AllowDrop = false;
                        DocCompareColorZone1.AllowDrop = false;
                        DocCompareSecondDocZone.AllowDrop = false;
                        DocCompareDragDropZone2.AllowDrop = false;
                        DocCompareColorZone2.AllowDrop = false;

                        Doc1StatsGrid.Visibility = Visibility.Collapsed;
                        Doc2StatsGrid.Visibility = Visibility.Collapsed;

                        Doc1PageNumberLabel.Content = "";
                        Doc2PageNumberLabel.Content = "";

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
                            DisplayImageLeft(docs.documentsToShow[0]);
                            ShowDoc1FileInfoButton.IsEnabled = true;
                            ShowDoc2FileInfoButton.IsEnabled = true;
                            OpenDoc1OriginalButton1.IsEnabled = true;
                            DocPreviewStatGrid.Visibility = Visibility.Visible;
                            Doc1StatsGrid.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            if (docs.documents.Count > 1)
                            {
                                docs.documentsToShow[0] = FindNextDocToShow();
                                DisplayImageLeft(docs.documentsToShow[0]);
                                OpenDoc1OriginalButton1.IsEnabled = true;
                                ShowDoc2FileInfoButton.IsEnabled = true;
                                ShowDoc1FileInfoButton.IsEnabled = true;
                                DocPreviewStatGrid.Visibility = Visibility.Visible;
                                Doc1StatsGrid.Visibility = Visibility.Visible;
                            }
                        }

                        if (docs.documents.Count > 1)
                        {
                            if (docs.documents[docs.documentsToShow[1]].processed == true)
                            {
                                DisplayImageMiddle(docs.documentsToShow[1]);
                                OpenDoc2OriginalButton2.IsEnabled = true;
                                ShowDoc2FileInfoButton.IsEnabled = true;
                                Doc2StatsGrid.Visibility = Visibility.Visible;
                            }
                            else
                            {
                                if (docs.documents.Count > 2)
                                {
                                    docs.documentsToShow[1] = FindNextDocToShow();
                                    DisplayImageMiddle(docs.documentsToShow[1]);
                                    OpenDoc2OriginalButton2.IsEnabled = true;
                                    ShowDoc2FileInfoButton.IsEnabled = true;
                                    Doc2StatsGrid.Visibility = Visibility.Visible;
                                }
                            }
                        }
                        else
                        {
                            Doc2StatsGrid.Visibility = Visibility.Collapsed;
                        }

                        if (docs.documents.Count > 2)
                        {
                            if (settings.numPanelsDragDrop == 3)
                            {
                                if (docs.documents[docs.documentsToShow[2]].processed == true)
                                {
                                    if (docs.documents.Count >= 3)
                                    {
                                        DisplayImageRight(docs.documentsToShow[2]);
                                        OpenDoc3OriginalButton3.IsEnabled = true;
                                    }
                                }
                                else
                                {
                                    if (docs.documents.Count > 3)
                                    {
                                        docs.documentsToShow[2] = FindNextDocToShow();
                                        DisplayImageRight(docs.documentsToShow[2]);
                                        OpenDoc3OriginalButton3.IsEnabled = true;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        HideDragDropZone2();
                        DocPreviewStatGrid.Visibility = Visibility.Collapsed;
                        Doc1StatsGrid.Visibility = Visibility.Visible;
                        Doc2StatsGrid.Visibility = Visibility.Collapsed;
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
                    }

                    ProgressBarDoc1.Visibility = Visibility.Hidden;
                    ProgressBarDoc2.Visibility = Visibility.Hidden;
                    ProgressBarDoc3.Visibility = Visibility.Hidden;

                    BrowseFileTopButton1.IsEnabled = true;
                    BrowseFileTopButton2.IsEnabled = true;
                    BrowseFileTopButton3.IsEnabled = true;

                    ReloadDoc1Button.IsEnabled = true;
                    ReloadDoc2Button.IsEnabled = true;
                    ReloadDoc3Button.IsEnabled = true;
                    BrowseFileButton1.IsEnabled = true;
                    BrowseFileButton2.IsEnabled = true;
                    BrowseFileButton3.IsEnabled = true;

                    DocCompareFirstDocZone.AllowDrop = true;
                    DocCompareDragDropZone1.AllowDrop = true;
                    DocCompareColorZone1.AllowDrop = true;
                    DocCompareSecondDocZone.AllowDrop = true;
                    DocCompareDragDropZone2.AllowDrop = true;
                    DocCompareColorZone2.AllowDrop = true;

                    CloseDoc1Button.IsEnabled = true;
                    CloseDoc2Button.IsEnabled = true;
                    CloseDoc3Button.IsEnabled = true;

                    OpenDoc1OriginalButton1.IsEnabled = true;
                    OpenDoc2OriginalButton2.IsEnabled = true;
                    OpenDoc3OriginalButton3.IsEnabled = true;

                    UpdateDocSelectionComboBox();

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
            docs.displayToReload = 3;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDocCompare2Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDocCompareReload.Visibility = Visibility.Visible;
            docs.forceAlignmentIndices = new List<List<int>>();
            docs.docToReload = docs.documentsToCompare[1];
            docs.displayToReload = 4;

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
            Dispatcher.Invoke(() => {
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
                            DisplayImageLeft(docs.documentsToShow[0]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(0);
                            break;

                        case 1:
                            DisplayImageMiddle(docs.documentsToShow[1]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(1);
                            break;

                        case 2:
                            DisplayImageRight(docs.documentsToShow[2]);
                            UpdateDocSelectionComboBox();
                            UpdateFileStat(2);
                            break;

                        case 3:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[0])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayImageLeft(docs.documentsToShow[0]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(1);
                                            break;

                                        case 1:
                                            DisplayImageMiddle(docs.documentsToShow[1]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(2);
                                            break;

                                        case 2:
                                            DisplayImageRight(docs.documentsToShow[2]);
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

                        case 4:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[1])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayImageLeft(docs.documentsToShow[0]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(0);
                                            break;

                                        case 1:
                                            DisplayImageMiddle(docs.documentsToShow[1]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(1);
                                            break;

                                        case 2:
                                            DisplayImageRight(docs.documentsToShow[2]);
                                            UpdateDocSelectionComboBox();
                                            UpdateFileStat(2);
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
                            DisplayImageLeft(docs.documentsToShow[0]);
                            break;

                        case 1:
                            DisplayImageMiddle(docs.documentsToShow[1]);
                            break;

                        case 2:
                            DisplayImageRight(docs.documentsToShow[2]);
                            break;

                        case 3:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[0])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayImageLeft(docs.documentsToShow[0]);
                                            break;

                                        case 1:
                                            DisplayImageMiddle(docs.documentsToShow[1]);
                                            break;

                                        case 2:
                                            DisplayImageRight(docs.documentsToShow[2]);
                                            break;
                                    }
                                }
                            }

                            ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                            break;

                        case 4:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[1])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            DisplayImageLeft(docs.documentsToShow[0]);
                                            break;

                                        case 1:
                                            DisplayImageMiddle(docs.documentsToShow[1]);
                                            break;

                                        case 2:
                                            DisplayImageRight(docs.documentsToShow[2]);
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
                docs.RemoveForceAligmentPairs(int.Parse(splittedName[1]));

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
                DisplayImageLeft(docs.documentsToShow[0]);
            if (docs.documents.Count >= 2)
                DisplayImageMiddle(docs.documentsToShow[1]);
            if (docs.documents.Count >= 2)
            {
                ShowDragDropZone3();
                Doc3Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone3.Visibility = Visibility.Visible;
            }

            if (docs.documents.Count >= 3)
                DisplayImageRight(docs.documentsToShow[2]);

            UpdateDocSelectionComboBox();
        }

        private void SettingsShowThirdPanelCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            settings.numPanelsDragDrop = 2;
            SaveSettings();
            docs.documentsToShow = new List<int>() { 0, 1 };
            DisplayImageLeft(docs.documentsToShow[0]);
            DisplayImageMiddle(docs.documentsToShow[1]);
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
            }
            else
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc1StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Collapsed;
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
            }
            else
            {
                Doc1StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc1StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc1StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc2StatCreatedLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc1StatLastEditorLabel0.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel.Visibility = Visibility.Collapsed;
                Doc2StatLastEditorLabel0.Visibility = Visibility.Collapsed;
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

            if (Doc3StatAuthorLabel0.Visibility == Visibility.Collapsed)
            {
                Doc3StatAuthorLabel0.Visibility = Visibility.Visible;
                Doc3StatAuthorLabel.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel0.Visibility = Visibility.Visible;
                Doc3StatCreatedLabel.Visibility = Visibility.Visible;
            }
            else
            {
                Doc3StatAuthorLabel0.Visibility = Visibility.Collapsed;
                Doc3StatAuthorLabel.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel0.Visibility = Visibility.Collapsed;
                Doc3StatCreatedLabel.Visibility = Visibility.Collapsed;
            }
        }

        private void ShowDocCompareFileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(3);

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
        }

        private void ShowDragDropZone3()
        {
            if (settings.numPanelsDragDrop == 3)
            {
                DocCompareThirdDocZone.Visibility = Visibility.Visible;
                DragDropPanel.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
            }
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

                                            foundButton.Visibility = Visibility.Visible;
                                        }
                                    }
                                }

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

                                            foundButton.Visibility = Visibility.Visible;
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

        private void SideGridMouseLeave(object sender, MouseEventArgs args)
        {
            Grid img = sender as Grid;
            string imgName = img.Tag.ToString();
            string[] splittedName;
            string nameToLook = "";
            string nameToLook2 = "";

            if (imgName.Contains("Left"))
            {
                splittedName = imgName.Split("Left");
                nameToLook = "SideButtonLeft" + splittedName[1];
                nameToLook2 = "SideButtonInvalidLeft" + splittedName[1];
            }

            if (imgName.Contains("Right"))
            {
                splittedName = imgName.Split("Right");
                nameToLook = "SideButtonRight" + splittedName[1];
                nameToLook2 = "SideButtonInvalidRight" + splittedName[1];
            }

            foreach (object child in img.Children)
            {
                if (child is Button)
                {
                    if ((child as Button).Tag.ToString() == nameToLook || (child as Button).Tag.ToString() == nameToLook2)
                    {
                        Button foundButton = child as Button;
                        if (inForceAlignMode == false)
                        {
                            foundButton.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            if (foundButton.Tag.ToString() == selectedSideGridButtonName1)
                            {
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

        private void SidePanelDocCompareButton_Click(object sender, RoutedEventArgs e)
        {
            if (docs.documents.Count >= 2 && docCompareRunning == false)
            {
                if (walkthroughMode == true && walkthroughStep == WalkthroughSteps.COMPARETAB)
                {
                    PopupCompareDocBubble.IsOpen = false;
                    walkthroughStep = WalkthroughSteps.COMPAREHIGHLIGHT; // 7
                }

                inForceAlignMode = false;
                //DocCompareLeftStatsGrid.Visibility = Visibility.Collapsed;
                //DocCompareRightStatsGrid.Visibility = Visibility.Collapsed;

                if (settings.isProVersion == true && settings.canSelectRefDoc == true)
                {
                    SetVisiblePanel(SidePanels.REFDOC);

                    // populate list box
                    ObservableCollection<string> items = new ObservableCollection<string>();
                    foreach (Document doc in docs.documents)
                    {
                        if (doc.processed == true)
                        {
                            items.Add(Path.GetFileName(doc.filePath));
                        }
                    }

                    RefDocListBox.ItemsSource = items;
                    RefDocListBox.SelectedIndex = 0;
                }
                else
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

            if (settings.numPanelsDragDrop == 3)
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
            }
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
                Doc2StatAuthorLabel.Text = docs.documents[docs.documentsToShow[2]].Creator;
                Doc2StatCreatedLabel.Text = docs.documents[docs.documentsToShow[2]].CreatedDate;
                Doc2StatLastEditorLabel.Text = docs.documents[docs.documentsToShow[2]].LastEditor;
                Doc2StatModifiedLabel.Text = docs.documents[docs.documentsToShow[2]].ModifiedDate;
            }

            if (i == 3) // DOC compare
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
            Doc1PageNumberLabel.Content = "";
            Doc2PageNumberLabel.Content = "";

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
            if(msgBox.ShowDialog() == true)
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
            foreach( DirectoryInfo di2 in di.EnumerateDirectories())
            {
                if(di2.Name != "lic")
                {
                    di2.Delete(true);
                }
            }

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