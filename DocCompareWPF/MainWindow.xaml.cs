using DocCompareWPF.Classes;
using DocCompareWPF.UIhelper;
using Microsoft.Win32;
using ProtoBuf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using ShellLink;

namespace DocCompareWPF
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string appDataDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare");
        private readonly DocumentManagement docs;
        private readonly string versionString = "1.3.2";
        private readonly string localetype = "DE";
        private readonly string workingDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare");
        private string compareResultFolder;
        private string compareResultFolder2;
        private bool docCompareRunning, docProcessRunning, animateDiffRunning;
        private MaskType showMask = 0;
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
        private Thread threadCompare, threadCompare2;
        private Thread threadAnimateDiff;
        private Thread threadDisplayResult;
        private Thread threadCheckTrial;
        private Thread threadCheckUpdate;
        private readonly Thread threadRenewLic;
        private Thread threadStartWalkthrough;

        // link scroll
        private bool[] linkScroll = { true, true, true, true, true };

        private WalkthroughSteps walkthroughStep = 0;

        public double HiddenPPTOpacity = 0.7;

        //private string FileFilter = FileFilter;
        private string FileFilter = "PPT files (*.pdf, *.ppt)|*.pdf;*.ppt;*.pptx;|All files|*.*";

        // mouse over hidden ppt slides effect buffer
        Effect hiddenPPTEffect;
        Visibility hiddenPPTVisi;

        SidePanels currentVisiblePanel;
        bool enableZoom;

        public MainWindow()
        {
            InitializeComponent();
            Directory.CreateDirectory(appDataDir);
            compareResultFolder = Path.Join(workingDir, Guid.NewGuid().ToString());
            compareResultFolder2 = Path.Join(workingDir, Guid.NewGuid().ToString());

            // GUI stuff
            showMask = MaskType.Magenta;
            AppVersionLabel.Content = "Version " + versionString;
            SetVisiblePanel(SidePanels.DRAGDROP);
            SidePanelDocCompareButton.IsEnabled = false;
            ActivateLicenseButton.IsEnabled = false;
            magnifier.Freeze(true);
            magnifier.ZoomFactor = 0.5;
            magnifier.Visibility = Visibility.Hidden;

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
                    numPanelsDragDrop = 2,
                    skipVersionString = versionString
                };

                SaveSettings();
            }

            if( localetype == "DE")
            {
                settings.cultureInfo = "de-de";
            }else
            {
                settings.cultureInfo = "en-us";
            }

            SaveSettings();

            docs = new DocumentManagement(settings.maxDocCount, workingDir, settings);

            // License Management
            try
            {
                LoadLicense();
                DisplayLicense();

                // history
                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
                {
                    lic.ConvertTrialToFree();
                    settings.maxDocCount = 2;
                    SaveSettings();
                    SaveLicense();
                    DisplayLicense();
                }
                else if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.ANNUAL_SUBSCRIPTION ||
                         lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.DEVELOPMENT)
                {
                    DisplayLicense();
                    settings.maxDocCount = 5;
                    SaveSettings();
                }

                ErrorHandling.ReportStatus("App Launch", "App version: " + versionString + " successfully launched with license: " + lic.GetLicenseTypesString() + ", expires/renewal on " + lic.GetExpiryDateString() + " on " + lic.GetUUID());
            }
            catch
            {
                lic = new LicenseManagement();
                lic.Init(); // init 7 days trial

                settings.maxDocCount = 5;
                SaveSettings();
                SaveLicense();

                CustomMessageBox msgBox = new CustomMessageBox();
                msgBox.Setup("2|Compare free", "You are using the free version of 2|Compare. If you like the functions, please consider making a subscription on www.hopie.tech.", "Okay");
                msgBox.ShowDialog();

                threadCheckTrial = new Thread(new ThreadStart(CheckTrial));
                threadCheckTrial.Start();

                ErrorHandling.ReportStatus("New free version", "on " + lic.GetUUID());
            }

            TimeSpan timeBuffer = lic.GetExpiryDate().Subtract(DateTime.Today);
            /*
            // Reminder to subscribe
            if (timeBuffer.TotalDays <= 5 && timeBuffer.TotalDays > 0 && lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
            {
                CustomMessageBox msgBox = new CustomMessageBox();
                msgBox.Setup("Expired lincense", "Your trial license will expire in " + timeBuffer.TotalDays + " day(s). Please consider making a subscription on www.hopietech.com", "Okay");
                msgBox.ShowDialog();
            }
            */

            // if license expires or needs renewal
            if (timeBuffer.TotalDays <= 0)
            {
                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.ANNUAL_SUBSCRIPTION)
                {
                    threadRenewLic = new Thread(new ThreadStart(RenewLicense));
                    threadRenewLic.Start();
                }
            }

            // Walkthrough if needed
            threadStartWalkthrough = new Thread(new ThreadStart(ShowWalkthroughStartMessage));
            threadStartWalkthrough.Start();

            // Check update if needed
            threadCheckUpdate = new Thread(new ThreadStart(CheckUpdate));
            threadCheckUpdate.Start();

            settings.FreeStartCount++;
            SaveSettings();

            if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
            {
                if (settings.FreeStartCount == 5)
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("2|Compare free", "You are using the free version of 2|Compare. If you like the functions, please consider making a subscription on https://hopie.tech", "Okay");
                    msgBox.ShowDialog();
                    settings.FreeStartCount = 0;
                    SaveSettings();
                }
            }

            EnableOpenOriginal();

        }

        private enum MaskType
        {
            Magenta,
            Green,
            Off,
        };

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

                            if (thisGrid.Tag != null)
                                if (thisGrid.Tag.ToString().Contains("Right") == false)
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

        private bool IsSupportedFile(string ext)
        {
            List<string> supportedFileExtensions = new List<string>
            {
                ".ppt",
                ".pptx",
                ".PPT",
                ".PPTX",
                ".pdf",
                ".PDF"
            };
            /*
            supportedFileExtensions.Add(".jpg");
            supportedFileExtensions.Add(".jpeg");
            supportedFileExtensions.Add(".JPG");
            supportedFileExtensions.Add(".JPEG");
            supportedFileExtensions.Add(".gif");
            supportedFileExtensions.Add(".GIF");
            supportedFileExtensions.Add(".png");
            supportedFileExtensions.Add(".PNG");
            supportedFileExtensions.Add(".bmp");
            supportedFileExtensions.Add(".BMP");
            */

            if (supportedFileExtensions.Contains(ext))
                return true;
            else
                return false;
        }

        private void BrowseFileButton1_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = FileFilter,
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
                    if (IsSupportedFile(ext) == false)
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

                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
                }
            }
        }

        private void BrowseFileButton2_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = FileFilter,
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
                    if (IsSupportedFile(ext) == false)
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
        }

        private void BrowseFileButton3_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = FileFilter,
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
                    if (IsSupportedFile(ext) == false)
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
            /*
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
            */

            Thread.Sleep(500);
            Dispatcher.Invoke(() => { DisplayLicense(); });
        }

        private async void CheckUpdate()
        {
            Thread.Sleep(1000);
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
                            msgBox.Setup("Update available", "A newer version of 2|Compare is available. Click OKAY to proceed with downloading the installer.\n\n" + res[2], "Okay", "Skip");

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
            catch
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

                    /*
                    BrowseFileButton1.IsEnabled = false;
                    DocCompareFirstDocZone.AllowDrop = false;
                    DocCompareDragDropZone1.AllowDrop = false;
                    DocCompareColorZone1.AllowDrop = false;
                    */
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
            DisableCloseDocument();

            // hide all then show them individually
            HideDragDropZone2();
            HideDragDropZone3();
            HideDragDropZone4();
            HideDragDropZone5();

            Doc2NameLabel.Content = "";
            Doc3NameLabel.Content = "";
            Doc4NameLabel.Content = "";
            Doc5NameLabel.Content = "";
            docs.doneGlobalAlignment = false;

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
                    Doc1Grid.Visibility = Visibility.Visible;
                    Doc1StatsGrid.Visibility = Visibility.Visible;
                    //ProgressBarDoc1.Visibility = Visibility.Visible;
                    Doc2Grid.Visibility = Visibility.Hidden;
                    Doc2StatsGrid.Visibility = Visibility.Collapsed;
                    UpdateDocSelectionComboBox();
                }

                if (docs.documents.Count >= 2)
                {
                    ShowDragDropZone3();
                    Doc2Grid.Visibility = Visibility.Visible;
                    Doc2StatsGrid.Visibility = Visibility.Visible;
                    //ProgressBarDoc2.Visibility = Visibility.Visible;
                    Doc3Grid.Visibility = Visibility.Hidden;
                    Doc3StatsGrid.Visibility = Visibility.Collapsed;
                    DragDropPanel.ColumnDefinitions[2].Width = new GridLength(0.7, GridUnitType.Star);
                    DocPreviewStatGrid.ColumnDefinitions[2].Width = new GridLength(0.7, GridUnitType.Star);

                }

                if (docs.documents.Count >= 3)
                {
                    ShowDragDropZone4();
                    Doc3Grid.Visibility = Visibility.Visible;
                    Doc3StatsGrid.Visibility = Visibility.Visible;
                    //ProgressBarDoc3.Visibility = Visibility.Visible;
                    Doc4Grid.Visibility = Visibility.Hidden;
                    Doc4StatsGrid.Visibility = Visibility.Collapsed;
                    DragDropPanel.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
                    DocPreviewStatGrid.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
                    DragDropPanel.ColumnDefinitions[3].Width = new GridLength(0.7, GridUnitType.Star);
                    DocPreviewStatGrid.ColumnDefinitions[3].Width = new GridLength(0.7, GridUnitType.Star);

                }

                if (docs.documents.Count >= 4)
                {
                    ShowDragDropZone5();
                    Doc4Grid.Visibility = Visibility.Visible;
                    Doc4StatsGrid.Visibility = Visibility.Visible;
                    //ProgressBarDoc4.Visibility = Visibility.Visible;
                    Doc5Grid.Visibility = Visibility.Hidden;
                    Doc5StatsGrid.Visibility = Visibility.Collapsed;
                    DragDropPanel.ColumnDefinitions[3].Width = new GridLength(1, GridUnitType.Star);
                    DocPreviewStatGrid.ColumnDefinitions[3].Width = new GridLength(1, GridUnitType.Star);
                    DragDropPanel.ColumnDefinitions[4].Width = new GridLength(0.7, GridUnitType.Star);
                    DocPreviewStatGrid.ColumnDefinitions[4].Width = new GridLength(0.7, GridUnitType.Star);
                }

                if (docs.documents.Count >= 5)
                {
                    DragDropPanel.ColumnDefinitions[4].Width = new GridLength(1, GridUnitType.Star);
                    DocPreviewStatGrid.ColumnDefinitions[4].Width = new GridLength(1, GridUnitType.Star);
                    //DisplayPreviewWithAligment(5, docs.documentsToShow[4], docs.globalAlignment);
                }

                threadCompare = new Thread(new ThreadStart(ComparePreviewThread));
                threadCompare.Start();

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
                showMask = MaskType.Magenta;
                Dispatcher.Invoke(() =>
                {
                    ShowMaskButtonMagenta.Visibility = Visibility.Hidden;
                    ShowMaskButtonGreen.Visibility = Visibility.Visible;
                    HideMaskButton.Visibility = Visibility.Hidden;
                    HighlightingDisableTip.Visibility = Visibility.Hidden;
                });

                int[,] forceIndices = new int[docs.forceAlignmentIndices.Count, 2];
                for (int i = 0; i < docs.forceAlignmentIndices.Count; i++)
                {
                    forceIndices[i, 0] = docs.forceAlignmentIndices[i][0];
                    forceIndices[i, 1] = docs.forceAlignmentIndices[i][1];
                }

                if (docs.forceAlignmentIndices.Count != 0)
                {
                    Dispatcher.Invoke(() =>
                    {
                        ManualAlignmentTip.Visibility = Visibility.Visible;
                    });
                }
                else
                {
                    Dispatcher.Invoke(() =>
                    {
                        ManualAlignmentTip.Visibility = Visibility.Hidden;
                    });
                }

                if (Directory.Exists(compareResultFolder))
                {
                    DirectoryInfo di = new DirectoryInfo(compareResultFolder);
                    di.Delete(true);
                }

                compareResultFolder = Path.Join(workingDir, Guid.NewGuid().ToString());
                Directory.CreateDirectory(compareResultFolder);

                if (Directory.Exists(compareResultFolder2))
                {
                    DirectoryInfo di = new DirectoryInfo(compareResultFolder2);
                    di.Delete(true);
                }

                compareResultFolder2 = Path.Join(workingDir, Guid.NewGuid().ToString());
                Directory.CreateDirectory(compareResultFolder2);

                Document.CompareDocs(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[1]].imageFolder, compareResultFolder, compareResultFolder2, out docs.pageCompareIndices, out docs.totalLen, forceIndices);
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

                        if (text1.Length != 0 || text2.Length != 0)
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
                });
            }
            catch
            {
                docCompareRunning = false;
            }
        }

        private void ComparePreviewThread()
        {
            docs.globalAlignment = null;
            Dispatcher.Invoke(() =>
            {
                ProgressBarGlobalAlignment.Visibility = Visibility.Visible;
            });

            ArrayList alignment = new ArrayList();
            string refDoc = docs.documents[docs.documentsToShow[0]].imageFolder;

            for (int i = 1; i < docs.documents.Count; i++)
            {
                alignment = DocCompareDLL.DocCompareClass.DocCompareMult(ref refDoc, ref docs.documents[docs.documentsToShow[i]].imageFolder, i - 1, alignment);
            }


            // inverse the sequence for showing

            for (int i = 0; i < alignment.Count; i++)
            {
                (alignment[i] as ArrayList).Reverse();
            }

            docs.globalAlignment = alignment;

            Dispatcher.Invoke(() =>
            {
                UpdateAllPreview();
                //UpdateDocSelectionComboBox();
                UpdateFileStat(0);

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
                EnableReload();
                EnableBrowseFile();

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

                EnableCloseDocument();
                EnableOpenOriginal();
            });


            docCompareRunning = false;
            docs.doneGlobalAlignment = true;
            Dispatcher.Invoke(() =>
            {
                ProgressBarGlobalAlignment.Visibility = Visibility.Hidden;
            });
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

        private string ReplaceInvalidXMLCharacters(string text)
        {
            string localText = text;
            
            localText = localText.Replace("&", "&#38;");            
            localText = localText.Replace(">", "&gt;");
            localText = localText.Replace("<", "&lt;");            
            localText = localText.Replace("\'", "&#39;");            
            localText = localText.Replace("\"", "&#34;");
            
            return localText;
        }

        private void DisplayComparisonResult()
        {
            Dispatcher.Invoke(() =>
            {
                DocCompareNameLabel1.Text = Path.GetFileName(docs.documents[docs.documentsToCompare[0]].filePath);
            });

            List<CompareMainItem> mainItemList = new List<CompareMainItem>();
            List<SideGridItemLeft> leftItemList = new List<SideGridItemLeft>();
            List<SideGridItemRight> rightItemList = new List<SideGridItemRight>();

            bool firstDiffFound = false;

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
                    ImgMaskRightName2 = "MainMaskImgRight2" + i.ToString(),
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
                    ShowPPTSpeakerNotesButtonNameRightChangedTrans = "ShowPPTSpeakerNotesButtonNameRightChanged" + i.ToString(),
                    ShowPPTSpeakerNotesButtonNameLeft = "ShowPPTSpeakerNotesButtonNameLeft" + i.ToString(),
                    ShowPPTSpeakerNotesButtonNameLeftChanged = "ShowPPTSpeakerNotesButtonNameLeftChanged" + i.ToString(),
                    PPTNoteGridLeftVisi = Visibility.Hidden,
                    PPTNoteGridRightVisi = Visibility.Hidden,
                    showPPTSpeakerNotesButtonRight = Visibility.Hidden,
                    showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden,
                    showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Hidden,
                    HiddenPPTOpacity = HiddenPPTOpacity,
                    AnimateButtonVisibility = Visibility.Hidden,
                };

                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                {
                    thisItem.PathToImgLeft = Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".png");

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
                            thisItem.PathToMaskImgRight2 = Path.Join(compareResultFolder2, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");

                            if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                            lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                            lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                            {
                                if (firstDiffFound == false)
                                {
                                    firstDiffFound = true;
                                    thisItem.AniDiffButtonEnable = true;
                                    thisItem.AniDiffButtonTooltip = "Click and hold to animate the difference";
                                }
                                else
                                {
                                    thisItem.AniDiffButtonEnable = true;
                                    thisItem.AniDiffButtonTooltip = "Animating difference is limited to the first page in the free version";
                                }
                            }
                            else
                            {
                                thisItem.AniDiffButtonEnable = true;
                                thisItem.AniDiffButtonTooltip = "Click and hold to animate the difference";
                            }

                            switch(showMask)
                            {
                                case MaskType.Magenta:
                                    thisItem.ShowMaskMagenta = Visibility.Visible;
                                    thisItem.ShowMaskGreen = Visibility.Hidden;
                                    break;
                                case MaskType.Green:
                                    thisItem.ShowMaskMagenta = Visibility.Hidden;
                                    thisItem.ShowMaskGreen = Visibility.Visible;
                                    break;
                                case MaskType.Off:
                                    thisItem.ShowMaskMagenta = Visibility.Hidden;
                                    thisItem.ShowMaskGreen = Visibility.Hidden;
                                    break;
                            }
                        }
                        else
                        {
                            if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                            lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                            lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                            {
                                thisItem.AniDiffButtonTooltip = "Comparison with hidden pages is only available in the pro version";
                            }
                            else
                            {
                                thisItem.AniDiffButtonTooltip = "The pages are identical";
                            }
                        }
                    }
                    else
                    {
                        if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                            lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                            lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                        {
                            thisItem.AniDiffButtonTooltip = "Animating difference is limited to the first page in the free version";
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

                bool showSpeakerNotesLeft = false;
                bool showSpeakerNotesLeftChanged = false;
                bool showSpeakerNotesRight = false;
                bool showSpeakerNotesRightChanged = false;
                bool showSpeakerNotesRightChangedTrans = false;
                bool didChange = false;

                if (docs.documents[docs.documentsToCompare[0]].fileType == Document.FileTypes.PPT && docs.documents[docs.documentsToCompare[1]].fileType == Document.FileTypes.PPT)
                {
                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1 && docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] == -1)
                    {
                        if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                        {
                            thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n<text>" + ReplaceInvalidXMLCharacters(docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]]) + "</text>";
                            showSpeakerNotesLeft = true;
                            showSpeakerNotesLeftChanged = false;
                        }
                        else
                        {
                            thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n" + "<text></text>";
                            showSpeakerNotesLeft = false;
                            showSpeakerNotesLeftChanged = false;
                        }

                    }
                    else if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] == -1 && docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                    {
                        if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                        {
                            thisItem.Document2 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n<text>" + ReplaceInvalidXMLCharacters(docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]]) + "</text>";
                            showSpeakerNotesRight = true;
                            showSpeakerNotesRightChanged = false;
                        }
                        else
                        {
                            thisItem.Document2 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n" + "<text></text>";
                            showSpeakerNotesRight = false;
                            showSpeakerNotesRightChanged = false;
                        }
                    }
                    else if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1 && docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                    {
                        string doc;


                        if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                        {
                            thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n<text>" + ReplaceInvalidXMLCharacters(docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]]) + "</text>";
                        }
                        else
                        {
                            thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n" + "<text></text>";
                        }

                        if (docs.pptSpeakerNotesDiff[i].Count != 0)
                        {
                            doc = "<?xml version =\"1.0\" encoding=\"UTF-8\"?> \n<text>";

                            foreach (DocCompareDLL.Diff diff in docs.pptSpeakerNotesDiff[i])
                            {
                                if (diff.operation == DocCompareDLL.Operation.INSERT)
                                {
                                    doc += "<INSERT>" + ReplaceInvalidXMLCharacters(diff.text) + "</INSERT>";
                                    didChange |= true;
                                }
                                else if (diff.operation == DocCompareDLL.Operation.DELETE)
                                {
                                    doc += "<DELETE>" + ReplaceInvalidXMLCharacters(diff.text) + "</DELETE>";
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

                        if (didChange == false)
                        {
                            if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length == 0
                                && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length == 0)
                            {
                                showSpeakerNotesLeft = false;
                                showSpeakerNotesLeftChanged = false;
                                showSpeakerNotesRight = false;
                                showSpeakerNotesRightChanged = false;
                            }
                            else
                            {
                                showSpeakerNotesLeft = true;
                                showSpeakerNotesLeftChanged = false;
                                showSpeakerNotesRight = true;
                                showSpeakerNotesRightChanged = false;
                            }
                        }
                        else
                        {
                            if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0
                                && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                            {
                                showSpeakerNotesLeft = true;
                                showSpeakerNotesLeftChanged = false;
                                showSpeakerNotesRight = false;
                                showSpeakerNotesRightChanged = true;
                            }
                            else if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length == 0
                                && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length == 0)
                            {
                                showSpeakerNotesLeft = false;
                                showSpeakerNotesLeftChanged = false;
                                showSpeakerNotesRight = false;
                                showSpeakerNotesRightChanged = false;
                            }
                            else if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length == 0
                                && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                            {
                                showSpeakerNotesLeft = false;
                                showSpeakerNotesLeftChanged = false;
                                showSpeakerNotesRight = false;
                                showSpeakerNotesRightChanged = true;
                            }
                            else if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0
                                && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length == 0)
                            {
                                showSpeakerNotesLeft = true;
                                showSpeakerNotesLeftChanged = false;
                                showSpeakerNotesRight = false;
                                showSpeakerNotesRightChanged = false;
                                showSpeakerNotesRightChangedTrans = true;
                            }
                            else
                            {
                                showSpeakerNotesLeft = false;
                                showSpeakerNotesLeftChanged = true;
                                showSpeakerNotesRight = false;
                                showSpeakerNotesRightChanged = true;
                            }
                        }

                    }
                }

                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                            lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                            lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                {
                    if (didChange == true)
                        thisItem.ShowSpeakerNotesTooltip = "The speaker notes was edited. Get the pro version to view the difference.";
                    else
                        thisItem.ShowSpeakerNotesTooltip = "Viewing of speaker notes is only available in the pro version.";
                }
                else
                {
                    if (didChange == true)
                        thisItem.ShowSpeakerNotesTooltip = "The speaker notes was edited. Click to view the difference.";
                    else
                        thisItem.ShowSpeakerNotesTooltip = "Click to show speaker notes";
                }

                if (showSpeakerNotesLeft)
                {
                    thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Visible;
                    thisItem.showPPTSpeakerNotesButtonLeftChanged = Visibility.Hidden;
                }
                else if (showSpeakerNotesLeftChanged)
                {
                    thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonLeftChanged = Visibility.Visible;
                }
                else
                {
                    thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonLeftChanged = Visibility.Hidden;
                }

                if (showSpeakerNotesRight)
                {
                    thisItem.showPPTSpeakerNotesButtonRight = Visibility.Visible;
                    thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Hidden;
                }
                else if (showSpeakerNotesRightChanged)
                {
                    thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Visible;
                    thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Hidden;
                }
                else if (showSpeakerNotesRightChangedTrans)
                {
                    thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Visible;
                }
                else
                {
                    thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                    thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Hidden;
                }

                // license

                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                    lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                    lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                {
                    thisItem.ShowSpeakerNoteEnable = true;
                    //thisItem.ShowSpeakerNotesTooltip = "Viewing of speaker notes is only available in the pro version";
                    thisItem.ShowHiddenEnable = Visibility.Visible;
                }
                else
                {
                    thisItem.ShowSpeakerNoteEnable = true;
                    //thisItem.ShowSpeakerNotesTooltip = "Click to show speaker notes";
                    thisItem.ShowHiddenEnable = Visibility.Hidden;
                }

                mainItemList.Add(thisItem);
            //}

            
            //for (int i = 0; i < docs.totalLen; i++)
            //{
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
                    RemoveForceAlignButtonName = "RemoveForceAlign" + i.ToString(),
                    HiddenPPTOpacity = HiddenPPTOpacity,
                };

                SideGridItemRight rightItem = new SideGridItemRight()
                {
                    GridName = "RightSideGrid" + i.ToString(),
                    ImgGridName = "SideImageRight" + i.ToString(),
                    ImgMaskName = "SideImageRightMask" + i.ToString(),
                    ImgMaskName2 = "SideImageRightMask2" + i.ToString(),
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
                    HiddenPPTOpacity = HiddenPPTOpacity,
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
                        rightItem.PathToMask2 = Path.Join(compareResultFolder2, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png");
                        if (i != 0)
                        {
                            leftItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            rightItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);


                            rightItem.DiffVisi = Visibility.Visible;
                            rightItem.NoDiffVisi = Visibility.Hidden;
                        }

                        switch (showMask)
                        {
                            case MaskType.Magenta:
                                rightItem.ShowMaskMagenta = Visibility.Visible;
                                rightItem.ShowMaskGreen = Visibility.Hidden;
                                break;
                            case MaskType.Green:
                                rightItem.ShowMaskMagenta = Visibility.Hidden;
                                rightItem.ShowMaskGreen = Visibility.Visible;
                                break;
                            case MaskType.Off:
                                rightItem.ShowMaskMagenta = Visibility.Hidden;
                                rightItem.ShowMaskGreen = Visibility.Hidden;
                                break;
                        }
                    }
                    else
                    {
                        if (didChange == false)
                        {
                            rightItem.DiffVisi = Visibility.Hidden;
                            rightItem.NoDiffVisi = Visibility.Visible;
                        }
                        else
                        {
                            leftItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            rightItem.BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            rightItem.DiffVisi = Visibility.Visible;
                            rightItem.NoDiffVisi = Visibility.Hidden;
                        }
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



                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                    lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                    lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                {
                    leftItem.LinkPagesToolTip = "Manual alignment is only available in the pro version";
                    rightItem.LinkPagesToolTip = "Manual alignment is only available in the pro version";
                    leftItem.ForceAlignEnable = true;
                    rightItem.ForceAlignEnable = true;

                    if (leftItem.ShowHidden == Visibility.Visible || rightItem.ShowHidden == Visibility.Visible)
                    {
                        rightItem.NoDiffVisi = Visibility.Hidden;
                        rightItem.DiffVisi = Visibility.Hidden;
                    }
                }
                else
                {
                    leftItem.LinkPagesToolTip = "Link pages";
                    rightItem.LinkPagesToolTip = "Link pages";
                    leftItem.ForceAlignEnable = true;
                    rightItem.ForceAlignEnable = true;
                }

                leftItemList.Add(leftItem);
                rightItemList.Add(rightItem);
            }

            Dispatcher.Invoke(() =>
            {
                DocCompareMainListView.ItemsSource = mainItemList;
                DocCompareSideListViewLeft.ItemsSource = leftItemList;
                DocCompareSideListViewRight.ItemsSource = rightItemList;
                //DocCompareSideScrollViewerLeft.Content = docCompareChildPanelLeft;
                //DocCompareSideScrollViewerRight.Content = docCompareChildPanelRight;

                docCompareGrid.Visibility = Visibility.Visible;
                ProgressBarDocCompare.Visibility = Visibility.Hidden;
                ProgressBarDocCompareAlign.Visibility = Visibility.Hidden;
                ProgressBarLoadingResults.Visibility = Visibility.Hidden;
            });
        }

        private void DisplayPreview(int viewerID, int docIndex)
        {
            if (docIndex != -1)
            {
                if (docs.documents.Count >= 1)
                {
                    List<SimpleImageItem> imageList = new List<SimpleImageItem>();

                    switch (viewerID)
                    {
                        case 1:
                            if(DocCompareListView1.Items.Count != 0)
                                imageList = DocCompareListView1.ItemsSource as List<SimpleImageItem>;
                            break;
                        case 2:
                            if (DocCompareListView2.Items.Count != 0)
                                imageList = DocCompareListView2.ItemsSource as List<SimpleImageItem>;
                            break;
                        case 3:
                            if (DocCompareListView3.Items.Count != 0)
                                imageList = DocCompareListView3.ItemsSource as List<SimpleImageItem>;
                            break;
                        case 4:
                            if (DocCompareListView4.Items.Count != 0)
                                imageList = DocCompareListView4.ItemsSource as List<SimpleImageItem>;
                            break;
                        case 5:
                            if (DocCompareListView5.Items.Count != 0)
                                imageList = DocCompareListView5.ItemsSource as List<SimpleImageItem>;
                            break;

                    }

                    if (docs.documents[docIndex].filePath != null)
                    {
                        int pageCounter = 0;

                        DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                        FileInfo[] fi = di.GetFiles();

                        if (fi.Length != 0)
                        {
                            for (int i = imageList.Count; i < fi.Length; i++)
                            {
                                try
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

                                    if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                                    lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                                    lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                                    {
                                        thisImage.ShowSpeakerNoteEnable = true;
                                        thisImage.ShowSpeakerNotesTooltip = "Viewing of speaker notes is only available in the pro version";
                                        thisImage.ShowHiddenEnable = Visibility.Visible;
                                    }
                                    else
                                    {
                                        thisImage.ShowSpeakerNoteEnable = true;
                                        thisImage.ShowSpeakerNotesTooltip = "Click to show speaker notes";
                                        thisImage.ShowHiddenEnable = Visibility.Hidden;
                                    }

                                    imageList.Add(thisImage);
                                    pageCounter++;
                                }catch
                                {
                                    // convert not yet complete
                                }
                            }
                        }

                        // add End of Document
                        if (imageList.Count == docs.documents[docIndex].totalCount - 1)
                        {
                            SimpleImageItem item = new SimpleImageItem
                            {
                                EoDVisi = Visibility.Visible,
                                showHidden = Visibility.Hidden,
                                showPPTSpeakerNotesButton = Visibility.Hidden,
                            };
                            imageList.Add(item);
                        }
                    }

                    if (imageList.Count != 0)
                    {
                        switch (viewerID)
                        {
                            case 1:
                                DocCompareListView1.ItemsSource = imageList;
                                DocCompareListView1.Items.Refresh();
                                //DocCompareListView1.ScrollIntoView(DocCompareListView1.Items[0]);
                                Doc1Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc1.Visibility = Visibility.Hidden;
                                //Doc1PageNumberLabel.Content = "1 / " + DocCompareListView1.Items.Count.ToString();
                                break;
                            case 2:
                                DocCompareListView2.ItemsSource = imageList;
                                DocCompareListView2.Items.Refresh();
                                //DocCompareListView2.ScrollIntoView(DocCompareListView2.Items[0]);
                                Doc2Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc2.Visibility = Visibility.Hidden;
                                Doc2NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[1]].filePath);
                                //Doc2PageNumberLabel.Content = "1 / " + DocCompareListView2.Items.Count.ToString();
                                break;
                            case 3:
                                DocCompareListView3.ItemsSource = imageList;
                                DocCompareListView3.Items.Refresh();
                                //DocCompareListView3.ScrollIntoView(DocCompareListView3.Items[0]);
                                Doc3Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc3.Visibility = Visibility.Hidden;
                                Doc3NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[2]].filePath);
                                //Doc3PageNumberLabel.Content = "1 / " + DocCompareListView3.Items.Count.ToString();
                                break;
                            case 4:
                                DocCompareListView4.ItemsSource = imageList;
                                DocCompareListView4.Items.Refresh();
                                //DocCompareListView4.ScrollIntoView(DocCompareListView4.Items[0]);
                                Doc4Grid.Visibility = Visibility.Visible;
                                ProgressBarDoc4.Visibility = Visibility.Hidden;
                                Doc4NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[3]].filePath);
                                //Doc4PageNumberLabel.Content = "1 / " + DocCompareListView4.Items.Count.ToString();
                                break;
                            case 5:
                                DocCompareListView5.ItemsSource = imageList;
                                DocCompareListView5.Items.Refresh();
                                //DocCompareListView5.ScrollIntoView(DocCompareListView5.Items[0]);
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

        private void DisplayPreviewWithAligment(int viewerID, int docIndex, ArrayList alignment)
        {
            int docIndexIntern = docs.documentsToShow.FindIndex(x => x == docIndex);

            if (docIndexIntern != -1)
            {
                if (docs.documents.Count >= 1)
                {
                    List<SimpleImageItem> imageList = new List<SimpleImageItem>();
                    if (docs.documents[docIndex].filePath != null)
                    {
                        int pageCounter = 0;

                        //loop here
                        for (int i = 0; i < (alignment[docIndexIntern] as ArrayList).Count; i++)
                        {
                            SimpleImageItem thisImage = new SimpleImageItem()
                            {
                                EoDVisi = Visibility.Hidden,
                                HiddenPPTOpacity = HiddenPPTOpacity,
                            };

                            if ((int)(alignment[docIndexIntern] as ArrayList)[i] != -1)
                            {
                                thisImage.PathToFile = Path.Join(docs.documents[docIndex].imageFolder, ((int)(alignment[docIndexIntern] as ArrayList)[i]).ToString() + ".png");

                                if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptIsHidden[(int)(alignment[docIndexIntern] as ArrayList)[i]] == true)
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

                                if (docs.documents[docIndex].fileType == Document.FileTypes.PPT && docs.documents[docIndex].pptSpeakerNotes[(int)(alignment[docIndexIntern] as ArrayList)[i]].Length != 0)
                                {
                                    thisImage.showPPTSpeakerNotesButton = Visibility.Visible;
                                    thisImage.ShowPPTSpeakerNotesButtonName = "ShowPPTNoteButton" + i.ToString();
                                    thisImage.PPTSpeakerNoteGridName = "PPTSpeakerNoteGrid" + i.ToString();
                                    thisImage.ClosePPTSpeakerNotesButtonName = "ClosePPTSpeakerNoteButton" + i.ToString();
                                    thisImage.PPTSpeakerNotes = docs.documents[docIndex].pptSpeakerNotes[(int)(alignment[docIndexIntern] as ArrayList)[i]];
                                }
                                else
                                {
                                    thisImage.showPPTSpeakerNotesButton = Visibility.Hidden;
                                }
                            }
                            else
                            {
                                thisImage.BlurRadius = 0;
                                thisImage.showHidden = Visibility.Hidden;
                                thisImage.showPPTSpeakerNotesButton = Visibility.Hidden;
                            }
                            // find the largest hidden image of same row

                            double WXHRatio = 100000;
                            int largestImgInd = docIndex;

                            for (int j = 0; j < alignment.Count; j++)
                            {
                                if ((int)(alignment[j] as ArrayList)[i] != -1)
                                {
                                    using (var imageStream = File.OpenRead(Path.Join(docs.documents[docs.documentsToShow[j]].imageFolder, (((int)(alignment[j] as ArrayList)[i])).ToString() + ".png")))
                                    {
                                        var decoder = BitmapDecoder.Create(imageStream, BitmapCreateOptions.IgnoreColorProfile,
                                            BitmapCacheOption.Default);
                                        var height1 = decoder.Frames[0].PixelHeight;
                                        var width1 = decoder.Frames[0].PixelWidth;
                                        double WXHLocal = (double)width1 / (double)height1;

                                        if (WXHRatio > WXHLocal)
                                        {
                                            WXHRatio = WXHLocal;
                                            largestImgInd = j;
                                        }
                                    }
                                }
                            }

                            if (docs.documentsToShow[largestImgInd] != docIndex)
                            {
                                thisImage.PathToFileHidden = Path.Join(docs.documents[docs.documentsToShow[largestImgInd]].imageFolder, (((int)(alignment[largestImgInd] as ArrayList)[i])).ToString() + ".png");
                            }

                            if (pageCounter == 0)
                            {
                                thisImage.Margin = new Thickness(10, 10, 10, 10);
                            }
                            else
                            {
                                thisImage.Margin = new Thickness(10, 0, 10, 10);
                            }


                            // license
                            if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                                lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                                lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                            {
                                thisImage.ShowSpeakerNoteEnable = true;
                                thisImage.ShowSpeakerNotesTooltip = "Viewing of speaker notes is only available in the pro version";
                                thisImage.ShowHiddenEnable = Visibility.Visible;
                            }
                            else
                            {
                                thisImage.ShowSpeakerNoteEnable = true;
                                thisImage.ShowSpeakerNotesTooltip = "Click to show speaker notes";
                                thisImage.ShowHiddenEnable = Visibility.Hidden;
                            }

                            imageList.Add(thisImage);
                            pageCounter++;

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
                    LicenseExpiryTypeLabel.Visibility = Visibility.Visible;
                    LicenseExpiryLabel.Visibility = Visibility.Visible;
                    LicenseStatusTypeLabel.Visibility = Visibility.Visible;
                    LicenseStatusLabel.Visibility = Visibility.Visible;
                    WindowGetProButton.Visibility = Visibility.Hidden;
                    (WindowGetProButton.Parent as Border).Visibility = Visibility.Hidden;
                    settings.maxDocCount = 5;
                    DocCompareDragDropZone3.IsEnabled = true;
                    BrowseFileButton3.IsEnabled = true;
                    DragDrop3ShowProVersion.Visibility = Visibility.Hidden;
                    DocCompareColorZone3.Visibility = Visibility.Visible;
                    HiddenPPTOpacity = 0.7;
                    EnableOpenOriginal();
                    EnableReload();
                    SaveSettings();
                    break;

                case LicenseManagement.LicenseTypes.TRIAL:
                case LicenseManagement.LicenseTypes.FREE:
                    LicenseTypeLabel.Content = "Free version";
                    LicenseExpiryTypeLabel.Content = "";
                    /*
                    TimeSpan timeBuffer = lic.GetExpiryDate().Subtract(DateTime.Today);
                    if (timeBuffer.TotalDays >= 0)
                        LicenseExpiryLabel.Content = timeBuffer.TotalDays.ToString() + " days";
                    else
                        LicenseExpiryTypeLabel.Content = "Expired";
                    */
                    LicenseExpiryLabel.Content = "";
                    LicenseExpiryTypeLabel.Visibility = Visibility.Collapsed;
                    LicenseExpiryLabel.Visibility = Visibility.Collapsed;
                    LicenseStatusTypeLabel.Visibility = Visibility.Collapsed;
                    LicenseStatusLabel.Visibility = Visibility.Collapsed;
                    WindowGetProButton.Visibility = Visibility.Visible;
                    (WindowGetProButton.Parent as Border).Visibility = Visibility.Visible;
                    settings.maxDocCount = 2;
                    DocCompareDragDropZone3.IsEnabled = false;
                    BrowseFileButton3.IsEnabled = false;
                    DragDrop3ShowProVersion.Visibility = Visibility.Visible;
                    DocCompareColorZone3.Visibility = Visibility.Hidden;
                    HiddenPPTOpacity = 1.0;
                    EnableOpenOriginal();
                    EnableReload();
                    SaveSettings();

                    break;

                case LicenseManagement.LicenseTypes.DEVELOPMENT:
                    LicenseTypeLabel.Content = "Developer license";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    LicenseExpiryLabel.Content = "- days";

                    ExtendTrialGrid1.Visibility = Visibility.Hidden;
                    ExtendTrialGrid2.Visibility = Visibility.Hidden;
                    LicenseExpiryTypeLabel.Visibility = Visibility.Collapsed;
                    LicenseExpiryLabel.Visibility = Visibility.Collapsed;
                    LicenseStatusTypeLabel.Visibility = Visibility.Collapsed;
                    LicenseStatusLabel.Visibility = Visibility.Collapsed;
                    WindowGetProButton.Visibility = Visibility.Hidden;
                    (WindowGetProButton.Parent as Border).Visibility = Visibility.Hidden;
                    settings.maxDocCount = 5;
                    DocCompareDragDropZone3.IsEnabled = true;
                    BrowseFileButton3.IsEnabled = true;
                    DragDrop3ShowProVersion.Visibility = Visibility.Hidden;
                    DocCompareColorZone3.Visibility = Visibility.Visible;
                    HiddenPPTOpacity = 0.7;
                    SaveSettings();
                    break;

                default:
                    LicenseTypeLabel.Content = "No license found";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    LicenseExpiryLabel.Content = "- days";
                    ExtendTrialGrid1.Visibility = Visibility.Hidden;
                    ExtendTrialGrid2.Visibility = Visibility.Hidden;
                    LicenseExpiryTypeLabel.Visibility = Visibility.Collapsed;
                    LicenseExpiryLabel.Visibility = Visibility.Collapsed;
                    LicenseStatusTypeLabel.Visibility = Visibility.Collapsed;
                    LicenseStatusLabel.Visibility = Visibility.Collapsed;
                    WindowGetProButton.Visibility = Visibility.Visible;
                    (WindowGetProButton.Parent as Border).Visibility = Visibility.Visible;
                    settings.maxDocCount = 2;
                    DocCompareDragDropZone3.IsEnabled = false;
                    BrowseFileButton3.IsEnabled = false;
                    DragDrop3ShowProVersion.Visibility = Visibility.Visible;
                    DocCompareColorZone3.Visibility = Visibility.Hidden;
                    HiddenPPTOpacity = 1.0;
                    SaveSettings();
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
                    DocCompareDragDropZone3.IsEnabled = true;
                    BrowseFileButton3.IsEnabled = true;
                    DragDrop3ShowProVersion.Visibility = Visibility.Hidden;
                    DocCompareColorZone3.Visibility = Visibility.Visible;
                    EnableOpenOriginal();
                    EnableReload();
                    break;

                case LicenseManagement.LicenseStatus.INACTIVE:
                    LicenseStatusTypeLabel.Content = "License status";
                    LicenseStatusLabel.Content = "Inactive";
                    DocCompareDragDropZone3.IsEnabled = false;
                    BrowseFileButton3.IsEnabled = false;
                    DragDrop3ShowProVersion.Visibility = Visibility.Visible;
                    DocCompareColorZone3.Visibility = Visibility.Hidden;
                    WindowGetProButton.Visibility = Visibility.Visible;
                    (WindowGetProButton.Parent as Border).Visibility = Visibility.Visible;
                    EnableOpenOriginal();
                    EnableReload();
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
                int previousSelection = docs.documentsToShow[0];

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

                //if(previousSelection != desiredInd)
                {
                    docs.doneGlobalAlignment = false;
                }

                //DisplayImageLeft(docs.documentsToShow[0]);
                /*
                if (docs.documents.Count >= 1)
                {
                    if (docs.documents.Count == 1)
                    {
                        DisplayPreview(1, docs.documentsToShow[0]);
                    }
                    else
                    {
                        DisplayPreviewWithAligment(1, docs.documentsToShow[0], docs.globalAlignment);
                    }
                }
                //DisplayPreview(1, docs.documentsToShow[0]);
                DisplayPreview(existingInd + 1, docs.documentsToShow[existingInd]);
                */

                if (docs.doneGlobalAlignment == false)
                {
                    docCompareRunning = true;
                    //ProgressBarDoc1.Visibility = Visibility.Visible;
                    threadCompare = new Thread(new ThreadStart(ComparePreviewThread));
                    threadCompare.Start();
                }
                else
                {
                    UpdateAllPreview();
                }

            }
            catch
            {
                Doc1NameLabelComboBox.SelectedIndex = 0;
                //UpdateDocSelectionComboBox();
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
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                string ext;

                foreach (string file in data)
                {
                    ext = Path.GetExtension(file);

                    if (ext == ".lnk")
                    {
                        Shortcut thisShortCut = Shortcut.ReadFromFile(file);
                        string linkedFile = thisShortCut.LinkTargetIDList.Path;
                        ext = Path.GetExtension(linkedFile);

                        if (IsSupportedFile(ext) == false)
                        {
                            ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(linkedFile));
                        }
                        else
                        {
                            if (docs.documents.Find(x => x.filePath == linkedFile) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                            {
                                docs.AddDocument(linkedFile);
                            }
                            else if (docs.documents.Count >= settings.maxDocCount)
                            {
                                ShowMaxDocCountWarningBox();
                                break;
                            }
                            else
                            {
                                ShowExistingDocCountWarningBox(linkedFile);
                            }
                        }
                    }
                    else
                    {
                        if (IsSupportedFile(ext) == false)
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

                    LoadFilesCommonPart();

                    threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                    threadLoadDocs.Start();

                    threadLoadDocsProgress = new Thread(new ThreadStart(ProcessDocProgressThread));
                    threadLoadDocsProgress.Start();
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

                    if (ext == ".lnk")
                    {
                        Shortcut thisShortCut = Shortcut.ReadFromFile(file);
                        string linkedFile = thisShortCut.LinkTargetIDList.Path;
                        ext = Path.GetExtension(linkedFile);

                        if (IsSupportedFile(ext) == false)
                        {
                            ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(linkedFile));
                        }
                        else
                        {
                            if (docs.documents.Find(x => x.filePath == linkedFile) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                            {
                                docs.AddDocument(linkedFile);
                            }
                            else if (docs.documents.Count >= settings.maxDocCount)
                            {
                                ShowMaxDocCountWarningBox();
                                break;
                            }
                            else
                            {
                                ShowExistingDocCountWarningBox(linkedFile);
                            }
                        }
                    }
                    else
                    {
                        if (IsSupportedFile(ext) == false)
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
            if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE &&
                            lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                            lic.GetLicenseStatus() != LicenseManagement.LicenseStatus.INACTIVE)
            {

                if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                    string ext;

                    foreach (string file in data)
                    {
                        ext = Path.GetExtension(file);

                        if (ext == ".lnk")
                        {
                            Shortcut thisShortCut = Shortcut.ReadFromFile(file);
                            string linkedFile = thisShortCut.LinkTargetIDList.Path;
                            ext = Path.GetExtension(linkedFile);

                            if (IsSupportedFile(ext) == false)
                            {
                                ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(linkedFile));
                            }
                            else
                            {
                                if (docs.documents.Find(x => x.filePath == linkedFile) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                                {
                                    docs.AddDocument(linkedFile);
                                }
                                else if (docs.documents.Count >= settings.maxDocCount)
                                {
                                    ShowMaxDocCountWarningBox();
                                    break;
                                }
                                else
                                {
                                    ShowExistingDocCountWarningBox(linkedFile);
                                }
                            }
                        }
                        else
                        {
                            if (IsSupportedFile(ext) == false)
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

                if (linkScroll[0] == true)
                {
                    Border border2;
                    if (linkScroll[1] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[2] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[3] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[4] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

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

                if (linkScroll[1] == true)
                {
                    Border border2;
                    if (linkScroll[0] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[2] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[3] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[4] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

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

                if (linkScroll[2] == true)
                {
                    Border border2;
                    if (linkScroll[0] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[1] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[3] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[4] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

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

                if (!(DocCompareMainListView.Items[int.Parse(splittedName[^1])] as CompareMainItem).AniDiffButtonTooltip.Contains("limited"))
                {

                    gridToAnimate = (sender as Button).Parent as Grid;
                    animateDiffRunning = true;

                    foreach (CompareMainItem item in DocCompareMainListView.Items)
                    {
                        if (item.AnimateDiffRightButtonName == (sender as Button).Tag.ToString())
                        {
                            switch (showMask)
                            {
                                case MaskType.Magenta:
                                    item.ShowMaskMagenta = Visibility.Visible;
                                    item.ShowMaskGreen = Visibility.Hidden;
                                    break;
                                case MaskType.Green:
                                    item.ShowMaskMagenta = Visibility.Hidden;
                                    item.ShowMaskGreen = Visibility.Visible;
                                    break;
                                case MaskType.Off:
                                    item.ShowMaskMagenta = Visibility.Hidden;
                                    item.ShowMaskGreen = Visibility.Hidden;
                                    break;
                            }
                            
                            item.BlurRadiusRight = 0;
                            item.ShowHiddenRight = Visibility.Hidden;
                        }
                    }

                    threadAnimateDiff = new Thread(new ThreadStart(AnimateDiffThread));
                    threadAnimateDiff.Start();
                }
                else
                {
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("2|Compare Pro", "Animating the difference is only available in the \n2|Compare Pro version. Visit www.hopie.tech for more information.");
                    msgBox.ShowDialog();
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

                switch (showMask)
                {
                    case MaskType.Magenta:
                        item.ShowMaskMagenta = Visibility.Visible;
                        item.ShowMaskGreen = Visibility.Hidden;
                        break;
                    case MaskType.Green:
                        item.ShowMaskMagenta = Visibility.Hidden;
                        item.ShowMaskGreen = Visibility.Visible;
                        break;
                    case MaskType.Off:
                        item.ShowMaskMagenta = Visibility.Hidden;
                        item.ShowMaskGreen = Visibility.Hidden;
                        break;
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

        }

        private void HandleMainDocCompareGridMouseEnter(object sender, MouseEventArgs args)
        {
            if (sender is Grid)
            {
                Grid parentGrid = sender as Grid;
                string name = parentGrid.Tag.ToString();
                string[] splitName;
                if (parentGrid.Tag.ToString().Contains("Left"))
                {
                    mainGridSelectedLeftOrRight = GridSelection.LEFT;
                    splitName = name.Split("Left");
                }
                else
                {
                    mainGridSelectedLeftOrRight = GridSelection.RIGHT;
                    splitName = name.Split("Right");
                }

                if ((DocCompareMainListView.Items[int.Parse(splitName[^1])] as CompareMainItem).AniDiffButtonEnable)
                {
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
                }

                Grid item = sender as Grid;
                Border childBorder = item.Children[0] as Border;
                Image img = childBorder.Child as Image;
                hiddenPPTEffect = img.Effect;
                img.Effect = null;


                if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                    lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE &&
                    lic.GetLicenseStatus() != LicenseManagement.LicenseStatus.INACTIVE)
                {
                    if (parentGrid.Tag.ToString().Contains("Left"))
                    {
                        hiddenPPTVisi = (item.Children[4] as Label).Visibility;
                        System.Windows.Shapes.Path path = item.Children[3] as System.Windows.Shapes.Path;
                        //path.Visibility = Visibility.Hidden;
                        Label label = item.Children[4] as Label;
                        //label.Visibility = Visibility.Hidden;
                        Grid grid = item.Children[2] as Grid;
                        grid.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        hiddenPPTVisi = (item.Children[6] as Label).Visibility;
                        System.Windows.Shapes.Path path = item.Children[5] as System.Windows.Shapes.Path;
                        //path.Visibility = Visibility.Hidden;
                        Label label = item.Children[6] as Label;
                        //label.Visibility = Visibility.Hidden;
                        Grid grid = item.Children[4] as Grid;
                        grid.Visibility = Visibility.Hidden;
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

                if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                    lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE &&
                    lic.GetLicenseStatus() != LicenseManagement.LicenseStatus.INACTIVE)
                {

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
                        System.Windows.Shapes.Path path = item.Children[5] as System.Windows.Shapes.Path;
                        path.Visibility = hiddenPPTVisi;
                        Label label = item.Children[6] as Label;
                        label.Visibility = hiddenPPTVisi;
                        Grid grid = item.Children[4] as Grid;
                        grid.Visibility = hiddenPPTVisi;
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
            showMask = MaskType.Off;

            foreach (CompareMainItem item in DocCompareMainListView.Items)
            {
                item.ShowMaskMagenta = Visibility.Hidden;
                item.ShowMaskGreen = Visibility.Hidden;
            }

            foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
            {
                item.ShowMaskMagenta = Visibility.Hidden;
                item.ShowMaskGreen = Visibility.Hidden;
            }

            ShowMaskButtonMagenta.Visibility = Visibility.Visible;
            ShowMaskButtonGreen.Visibility = Visibility.Hidden;
            HideMaskButton.Visibility = Visibility.Hidden;
            HighlightingDisableTip.Visibility = Visibility.Visible;
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

                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1 && docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                        {
                            bool didChange = false;  
                            if (docs.pptSpeakerNotesDiff[i].Count != 0)
                            {
                                foreach (DocCompareDLL.Diff diff in docs.pptSpeakerNotesDiff[i])
                                {
                                    if (diff.operation == DocCompareDLL.Operation.INSERT)
                                    {
                                       didChange |= true;
                                    }
                                    else if (diff.operation == DocCompareDLL.Operation.DELETE)
                                    {
                                        didChange |= true;
                                    }
                                }
                            }

                            if(didChange == true)
                            {
                                (DocCompareSideListViewLeft.Items[i] as SideGridItemLeft).BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                                (DocCompareSideListViewRight.Items[i] as SideGridItemRight).BackgroundBrush = Color.FromArgb(128, 255, 44, 108);
                            }
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
            docs.doneGlobalAlignment = false;

            if (docs.documents.Count >= 1)
            {
                if (docs.documents[docs.documentsToShow[0]].processed == false)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc1Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc1.Visibility = Visibility.Visible;
                        ShowDragDropZone2();
                        DocCompareDragDropZone2.IsEnabled = true;
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
                        DragDropPanel.ColumnDefinitions[2].Width = new GridLength(0.5, GridUnitType.Star);
                        DocPreviewStatGrid.ColumnDefinitions[2].Width = new GridLength(0.5, GridUnitType.Star);
                        DragDrop3ShowProVersion.Visibility = Visibility.Hidden;
                        DocCompareColorZone3.Visibility = Visibility.Visible;
                        DocCompareDragDropZone3.IsEnabled = true;
                    });
                }

                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                   lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                   lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                {
                    //DocCompareDragDropZone3.IsEnabled = false;
                    BrowseFileButton3.IsEnabled = false;
                    DragDrop3ShowProVersion.Visibility = Visibility.Visible;
                    DocCompareColorZone3.Visibility = Visibility.Hidden;
                    DocCompareDragDropZone3.IsEnabled = false;
                    //HideDragDropZone3();
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
                        DragDropPanel.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
                        DragDropPanel.ColumnDefinitions[3].Width = new GridLength(0.5, GridUnitType.Star);
                        DocPreviewStatGrid.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
                        DocPreviewStatGrid.ColumnDefinitions[3].Width = new GridLength(0.5, GridUnitType.Star);
                        DocCompareDragDropZone4.IsEnabled = true;
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
                        DragDropPanel.ColumnDefinitions[3].Width = new GridLength(1, GridUnitType.Star);
                        DragDropPanel.ColumnDefinitions[4].Width = new GridLength(0.5, GridUnitType.Star);
                        DocPreviewStatGrid.ColumnDefinitions[3].Width = new GridLength(1, GridUnitType.Star);
                        DocPreviewStatGrid.ColumnDefinitions[4].Width = new GridLength(0.5, GridUnitType.Star);
                        DocCompareDragDropZone5.IsEnabled = true;
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
                        DragDropPanel.ColumnDefinitions[4].Width = new GridLength(1, GridUnitType.Star);
                        DocPreviewStatGrid.ColumnDefinitions[4].Width = new GridLength(1, GridUnitType.Star);
                    });
                }
            }

            docProcessRunning = true;

            Application.Current.Dispatcher.Invoke(() =>
            {
                ProcessingDocProgressCard.Visibility = Visibility.Visible;
                ProcessingDocProgressbar.Value = 0;
                ProcessingDocLabel.Text = "Processing: " + Path.GetFileName(docs.documents[0].filePath);

                DisableReload();
                DisableCloseDocument();
                DisableOpenOriginal();
                DisableShowDocInfo();

                // clear all view
                DocCompareListView1.ItemsSource = new List<SimpleImageItem>();
                DocCompareListView2.ItemsSource = new List<SimpleImageItem>();
                DocCompareListView3.ItemsSource = new List<SimpleImageItem>();
                DocCompareListView4.ItemsSource = new List<SimpleImageItem>();
                DocCompareListView5.ItemsSource = new List<SimpleImageItem>();
                DocCompareListView1.Items.Refresh();
                DocCompareListView2.Items.Refresh();
                DocCompareListView3.Items.Refresh();
                DocCompareListView4.Items.Refresh();
                DocCompareListView5.Items.Refresh();
                
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

        private void UpdateAllPreview()
        {
            if (docs.documents.Count != 0)
            {
                if (docs.documents.Count > 0)
                {
                    if (docs.documents[docs.documentsToShow[0]].readyToShow == true)
                    {
                        ProgressBarDoc1.Visibility = Visibility.Visible;

                        if (docs.documents.Count == 1)
                        {
                            DisplayPreview(1, docs.documentsToShow[0]);
                        }
                        else
                        {
                            if (docs.globalAlignment != null && docs.globalAlignment.Count != 0)
                                DisplayPreviewWithAligment(1, docs.documentsToShow[0], docs.globalAlignment);
                            else
                                DisplayPreview(1, docs.documentsToShow[0]);
                        }

                        ShowDoc1FileInfoButton.IsEnabled = true;
                        ShowDoc2FileInfoButton.IsEnabled = true;
                        //OpenDoc1OriginalButton1.IsEnabled = true;
                        DocPreviewStatGrid.Visibility = Visibility.Visible;
                        Doc1StatsGrid.Visibility = Visibility.Visible;
                        UpdateFileStat(0);
                        ProgressBarDoc1.Visibility = Visibility.Hidden;
                        DocCompareColorZone1.Visibility = Visibility.Visible;
                    }
                }

                if (docs.documents.Count > 1)
                {
                    if (docs.documents[docs.documentsToShow[1]].readyToShow == true)
                    {
                        ProgressBarDoc2.Visibility = Visibility.Visible;
                        if (docs.globalAlignment != null && docs.globalAlignment.Count != 0)
                            DisplayPreviewWithAligment(2, docs.documentsToShow[1], docs.globalAlignment);
                        else
                            DisplayPreview(2, docs.documentsToShow[1]);
                        //DisplayPreview(2, docs.documentsToShow[1]);
                        //OpenDoc2OriginalButton2.IsEnabled = true;
                        ShowDoc2FileInfoButton.IsEnabled = true;
                        Doc2StatsGrid.Visibility = Visibility.Visible;
                        Doc2NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[1]].filePath);
                        UpdateFileStat(1);
                        ProgressBarDoc2.Visibility = Visibility.Hidden;
                        DocCompareColorZone2.Visibility = Visibility.Visible;
                    }

                }

                if (docs.documents.Count > 2)
                {
                    // TODO: Premium
                    //if (settings.numPanelsDragDrop == 3)
                    {
                        if (docs.documents[docs.documentsToShow[2]].readyToShow == true)
                        {
                            ProgressBarDoc3.Visibility = Visibility.Visible;
                            if (docs.globalAlignment != null && docs.globalAlignment.Count != 0)
                                DisplayPreviewWithAligment(3, docs.documentsToShow[2], docs.globalAlignment);
                            else
                                DisplayPreview(3, docs.documentsToShow[2]);
                            //DisplayPreview(3, docs.documentsToShow[2]);
                            //OpenDoc3OriginalButton3.IsEnabled = true;
                            ShowDoc3FileInfoButton.IsEnabled = true;
                            Doc3StatsGrid.Visibility = Visibility.Visible;
                            Doc3NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[2]].filePath);
                            UpdateFileStat(2);
                            ProgressBarDoc3.Visibility = Visibility.Hidden;
                            DocCompareColorZone3.Visibility = Visibility.Visible;
                        }
                    }
                }

                if (docs.documents.Count > 3)
                {
                    // TODO: Premium
                    //if (settings.numPanelsDragDrop == 3)
                    {
                        if (docs.documents[docs.documentsToShow[3]].readyToShow == true)
                        {
                            ProgressBarDoc4.Visibility = Visibility.Visible;
                            if (docs.globalAlignment != null && docs.globalAlignment.Count != 0)
                                DisplayPreviewWithAligment(4, docs.documentsToShow[3], docs.globalAlignment);
                            else
                                DisplayPreview(4, docs.documentsToShow[3]);
                            //DisplayPreview(4, docs.documentsToShow[3]);
                            //OpenDoc4OriginalButton4.IsEnabled = true;
                            ShowDoc4FileInfoButton.IsEnabled = true;
                            Doc4StatsGrid.Visibility = Visibility.Visible;
                            Doc4NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[3]].filePath);
                            UpdateFileStat(3);
                            ProgressBarDoc4.Visibility = Visibility.Hidden;
                            DocCompareColorZone4.Visibility = Visibility.Visible;
                        }
                    }
                }

                if (docs.documents.Count > 4)
                {
                    // TODO: Premium
                    //if (settings.numPanelsDragDrop == 3)
                    {
                        if (docs.documents[docs.documentsToShow[4]].readyToShow == true)
                        {
                            ProgressBarDoc5.Visibility = Visibility.Visible;
                            if (docs.globalAlignment != null && docs.globalAlignment.Count != 0)
                                DisplayPreviewWithAligment(5, docs.documentsToShow[4], docs.globalAlignment);
                            else
                                DisplayPreview(5, docs.documentsToShow[4]);
                            //DisplayPreview(5, docs.documentsToShow[4]);
                            //OpenDoc5OriginalButton5.IsEnabled = true;
                            ShowDoc5FileInfoButton.IsEnabled = true;
                            Doc5StatsGrid.Visibility = Visibility.Visible;
                            Doc5NameLabel.Content = Path.GetFileName(docs.documents[docs.documentsToShow[4]].filePath);
                            UpdateFileStat(4);
                            ProgressBarDoc5.Visibility = Visibility.Hidden;
                            DocCompareColorZone5.Visibility = Visibility.Visible;
                        }
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
        }

        private void ProcessDocProgressThread()
        {
            Dispatcher.Invoke(() =>
            {
                DisableBrowseFile();

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

                //if (docs.documents.Count >= 1)
                    DocCompareColorZone1.Visibility = Visibility.Hidden;

                //if (docs.documents.Count >= 2)
                    DocCompareColorZone2.Visibility = Visibility.Hidden;

                //if (docs.documents.Count >= 3)
                    DocCompareColorZone3.Visibility = Visibility.Hidden;

                //if (docs.documents.Count >= 4)
                    DocCompareColorZone4.Visibility = Visibility.Hidden;

                //if (docs.documents.Count >= 5)
                    DocCompareColorZone5.Visibility = Visibility.Hidden;

                ProcessingDocProgressbar.Value = 0.0;

            });

            try
            {
                while (docProcessRunning == true)
                {
                    try
                    {
                        Dispatcher.Invoke(() =>
                        {
                            if (docProcessingCounter < docs.documents.Count)
                            {
                                ProcessingDocProgressbar.Value = (((double)docs.documents[docProcessingCounter].currIndex / (double)docs.documents[docProcessingCounter].totalCount) * (1.0 / (double)docs.documents.Count)  + ((double)docProcessingCounter) / (double)docs.documents.Count ) * 100.0;

                                ProcessingDocLabel.Text = "Processing: " + Path.GetFileName(docs.documents[docProcessingCounter].filePath) + " (" + docs.documents[docProcessingCounter].currIndex.ToString() + "/" + (docs.documents[docProcessingCounter].totalCount - 1).ToString() + ")";
                                if (docs.documents[docProcessingCounter].currIndex % 10 == 0 || docs.documents[docProcessingCounter].currIndex == docs.documents[docProcessingCounter].totalCount - 1)
                                {
                                    UpdateAllPreview();
                                }
                            }

                        });
                    }
                    catch
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBarDoc1.Visibility = Visibility.Hidden;
                            ProgressBarDoc2.Visibility = Visibility.Hidden;
                            ProgressBarDoc3.Visibility = Visibility.Hidden;
                            ProgressBarDoc4.Visibility = Visibility.Hidden;
                            ProgressBarDoc5.Visibility = Visibility.Hidden;

                            DocCompareColorZone1.Visibility = Visibility.Visible;
                            DocCompareColorZone2.Visibility = Visibility.Visible;
                            DocCompareColorZone3.Visibility = Visibility.Visible;
                            DocCompareColorZone4.Visibility = Visibility.Visible;
                            DocCompareColorZone5.Visibility = Visibility.Visible;

                            EnableBrowseFile();

                        });
                    }

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


                    UpdateDocSelectionComboBox();

                    if (docs.documents.Count > 1)
                    {
                        if (docs.doneGlobalAlignment == false)
                        {
                            threadCompare = new Thread(new ThreadStart(ComparePreviewThread));
                            threadCompare.Start();
                        }
                    }
                    else
                    {
                        UpdateAllPreview();
                    }

                    ShowInfoButtonSetVisi();

                    if (docs.documents.Count >= 2)
                        SidePanelDocCompareButton.IsEnabled = true;

                    ProcessingDocProgressCard.Visibility = Visibility.Hidden;

                    if (docs.documents.Count >= 1)
                    {
                        DocCompareColorZone2.Visibility = Visibility.Visible;
                        DocCompareColorZone3.Visibility = Visibility.Visible;
                        DocCompareColorZone4.Visibility = Visibility.Visible;
                        DocCompareColorZone5.Visibility = Visibility.Visible;                        
                    }
                    else if (docs.documents.Count >= 2)
                    {
                        DocCompareColorZone3.Visibility = Visibility.Visible;
                        DocCompareColorZone4.Visibility = Visibility.Visible;
                        DocCompareColorZone5.Visibility = Visibility.Visible;
                    }
                    else if (docs.documents.Count >= 3)
                    {
                        DocCompareColorZone4.Visibility = Visibility.Visible;
                        DocCompareColorZone5.Visibility = Visibility.Visible;
                    }
                    else if (docs.documents.Count >= 4)
                    {
                        DocCompareColorZone5.Visibility = Visibility.Visible;
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
                        Doc1NameLabelComboBox.Margin = new Thickness(0, 0, 30, 0);
                        Doc2NameLabel.Margin = new Thickness(0);
                        Doc3NameLabel.Margin = new Thickness(0);
                        Doc4NameLabel.Margin = new Thickness(0);
                        Doc5NameLabel.Margin = new Thickness(0);
                        break;
                    case 2:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        Doc1NameLabelComboBox.Margin = new Thickness(0, 0, 0, 0);
                        Doc2NameLabel.Margin = new Thickness(0, 0, 30, 0);
                        Doc3NameLabel.Margin = new Thickness(0);
                        Doc4NameLabel.Margin = new Thickness(0);
                        Doc5NameLabel.Margin = new Thickness(0);
                        break;
                    case 3:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Visible;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        Doc1NameLabelComboBox.Margin = new Thickness(0, 0, 0, 0);
                        Doc2NameLabel.Margin = new Thickness(0);
                        Doc3NameLabel.Margin = new Thickness(0, 0, 30, 0);
                        Doc4NameLabel.Margin = new Thickness(0);
                        Doc5NameLabel.Margin = new Thickness(0);
                        break;
                    case 4:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Visible;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        Doc1NameLabelComboBox.Margin = new Thickness(0, 0, 0, 0);
                        Doc2NameLabel.Margin = new Thickness(0);
                        Doc3NameLabel.Margin = new Thickness(0);
                        Doc4NameLabel.Margin = new Thickness(0, 0, 30, 0);
                        Doc5NameLabel.Margin = new Thickness(0);
                        break;
                    case 5:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Visible;
                        Doc1NameLabelComboBox.Margin = new Thickness(0, 0, 0, 0);
                        Doc2NameLabel.Margin = new Thickness(0);
                        Doc3NameLabel.Margin = new Thickness(0);
                        Doc4NameLabel.Margin = new Thickness(0);
                        Doc5NameLabel.Margin = new Thickness(0, 0, 30, 0);
                        break;
                    default:
                        ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
                        ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
                        Doc1NameLabelComboBox.Margin = new Thickness(0, 0, 0, 0);
                        Doc2NameLabel.Margin = new Thickness(0);
                        Doc3NameLabel.Margin = new Thickness(0);
                        Doc4NameLabel.Margin = new Thickness(0);
                        Doc5NameLabel.Margin = new Thickness(0);
                        break;
                }
            });
        }

        private void ProcessDocThread()
        {
            // Going through documents in stack, check if reloading needed
            docProcessingCounter = 0;
            docs.doneGlobalAlignment = false;

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
                                docs.documents[i].readyToShow = true;
                                break;

                            case Document.FileTypes.PPT:
                                ret = docs.documents[i].ReadPPT();
                                break;

                            case Document.FileTypes.PIC:
                                ret = docs.documents[i].ReadPic();
                                docs.documents[i].readyToShow = true;
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
                        else if (ret == -3)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                CustomMessageBox msgBox = new CustomMessageBox();
                                msgBox.Setup("Empty PowerPoint file", "The selected file " + Path.GetFileName(docs.documents[i].filePath) + " is an empty PowerPoint file.", "Okay");
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
            docCompareGrid.Visibility = Visibility.Hidden;
            ProgressBarDocCompareReload.Visibility = Visibility.Visible;
            docs.forceAlignmentIndices = new List<List<int>>();

            docs.docToReload = docs.documentsToCompare[0];
            docs.displayToReload = 5;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDocCompare2Button_Click(object sender, RoutedEventArgs e)
        {
            docCompareGrid.Visibility = Visibility.Hidden;
            ProgressBarDocCompareReload.Visibility = Visibility.Visible;
            docs.forceAlignmentIndices = new List<List<int>>();
            docs.docToReload = docs.documentsToCompare[1];
            docs.displayToReload = 6;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDocThread()
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBarDocCompareReload.Visibility = Visibility.Visible;
                DisableReload();
            });

            if (docs.documents[docs.docToReload].ReloadDocument(workingDir) == 0)
            {
                docs.documents[docs.docToReload].ReadStats(settings.cultureInfo);

                Dispatcher.Invoke(() =>
                {
                    switch (docs.displayToReload)
                    {
                        case 0:
                        case 1:
                        case 2:
                        case 3:
                        case 4:
                            docCompareRunning = true;
                            //ProgressBarDoc1.Visibility = Visibility.Visible;
                            threadCompare = new Thread(new ThreadStart(ComparePreviewThread));
                            threadCompare.Start();
                            break;

                        case 5:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[0])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                        case 1:
                                        case 2:
                                        case 3:
                                        case 4:
                                            docCompareRunning = true;
                                            //ProgressBarDoc1.Visibility = Visibility.Visible;
                                            threadCompare = new Thread(new ThreadStart(ComparePreviewThread));
                                            threadCompare.Start();
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
                            threadCompare2 = new Thread(new ThreadStart(CompareDocsThread));
                            threadCompare2.Start();
                            break;

                        case 6:
                            for (int i = 0; i < docs.documentsToShow.Count; i++)
                            {
                                if (docs.documentsToShow[i] == docs.documentsToCompare[1])
                                {
                                    switch (i)
                                    {
                                        case 0:
                                        case 1:
                                        case 2:
                                        case 3:
                                        case 4:
                                            docCompareRunning = true;
                                            //ProgressBarDoc1.Visibility = Visibility.Visible;
                                            threadCompare = new Thread(new ThreadStart(ComparePreviewThread));
                                            threadCompare.Start();
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
                            threadCompare2 = new Thread(new ThreadStart(CompareDocsThread));
                            threadCompare2.Start();
                            break;
                    }

                    if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                    lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE)
                    {
                        ReloadDoc1Button.IsEnabled = false;
                        ReloadDoc2Button.IsEnabled = false;
                    }

                    EnableReload();
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
                            /*
                            BrowseFileButton1.IsEnabled = false;
                            DocCompareFirstDocZone.AllowDrop = false;
                            DocCompareDragDropZone1.AllowDrop = false;
                            DocCompareColorZone1.AllowDrop = false;
                            */
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

                    if (enableZoom == true)
                    {
                        magnifier.Visibility = Visibility.Hidden;
                        magnifier.ZoomFactor = 1.0;
                        magnifier.Freeze(true);
                        enableZoom = false;
                        EnableZoomButton.Visibility = Visibility.Visible;
                        DisableZoomButton.Visibility = Visibility.Hidden;
                        ZoomButtonBackground2.Visibility = Visibility.Hidden;
                    }
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
                    if (enableZoom == true)
                    {
                        enableZoom = false;
                        EnableZoomButton.Visibility = Visibility.Visible;
                        DisableZoomButton.Visibility = Visibility.Hidden;
                        ZoomButtonBackground2.Visibility = Visibility.Hidden;
                    }
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

            ShowDoc1FileInfoButton.Visibility = Visibility.Hidden;
            HideDoc1FileInfoButton.Visibility = Visibility.Visible;

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

            ShowDoc2FileInfoButton.Visibility = Visibility.Hidden;
            HideDoc2FileInfoButton.Visibility = Visibility.Visible;

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

        private void ShowDoc3FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(2);

            ShowDoc3FileInfoButton.Visibility = Visibility.Hidden;
            HideDoc3FileInfoButton.Visibility = Visibility.Visible;


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

        private void ShowDocCompareFileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(5);
            ShowDocCompareFileInfoButton.Visibility = Visibility.Hidden;
            HideDocCompareFileInfoButton.Visibility = Visibility.Visible;

            DocCompareLeftStatAuthorLabel.Visibility = Visibility.Visible;
            DocCompareLeftStatCreatedLabel.Visibility = Visibility.Visible;
            DocCompareLeftStatAuthorLabel0.Visibility = Visibility.Visible;
            DocCompareLeftStatCreatedLabel0.Visibility = Visibility.Visible;
            DocCompareLeftStatLastEditorLabel0.Visibility = Visibility.Visible;
            DocCompareLeftStatLastEditorLabel.Visibility = Visibility.Visible;

            DocCompareRightStatAuthorLabel.Visibility = Visibility.Visible;
            DocCompareRightStatCreatedLabel.Visibility = Visibility.Visible;
            DocCompareRightStatAuthorLabel0.Visibility = Visibility.Visible;
            DocCompareRightStatCreatedLabel0.Visibility = Visibility.Visible;
            DocCompareRightStatLastEditorLabel0.Visibility = Visibility.Visible;
            DocCompareRightStatLastEditorLabel.Visibility = Visibility.Visible;

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
            showMask = MaskType.Magenta;

            foreach (CompareMainItem item in DocCompareMainListView.Items)
            {
                item.ShowMaskMagenta = Visibility.Visible;
                item.ShowMaskGreen = Visibility.Hidden;
            }

            foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
            {
                item.ShowMaskMagenta = Visibility.Visible;
                item.ShowMaskGreen = Visibility.Hidden;
            }            

            ShowMaskButtonMagenta.Visibility = Visibility.Hidden;
            ShowMaskButtonGreen.Visibility = Visibility.Visible;
            HideMaskButton.Visibility = Visibility.Hidden;
            HighlightingDisableTip.Visibility = Visibility.Hidden;
        }

        private void ShowMaxDocCountWarningBox()
        {
            CustomMessageBox msgBox = new CustomMessageBox();

            if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
            {
                msgBox.Setup("Maximum number of documents loaded", "You have selected more than " + settings.maxDocCount.ToString() + " documents. Please consider subscribing the pro version on www.hopie.tech to load up to 5 documents.", "Okay");
            }
            else
            {
                msgBox.Setup("Maximum number of documents loaded", "You have selected more than " + settings.maxDocCount.ToString() + " documents. Only the first " + settings.maxDocCount.ToString() + " documents are loaded.", "Okay");
            }
            msgBox.ShowDialog();
        }

        private void SideGridButtonMouseClick(object sender, RoutedEventArgs args)
        {
            Border border = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewLeft, 0);
            ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;
            Border border2 = (Border)VisualTreeHelper.GetChild(DocCompareSideListViewRight, 0);
            ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

            Button button = sender as Button;

            if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
            {
                CustomMessageBox msgBox = new CustomMessageBox();
                msgBox.Setup("2|Compare Pro", "Manual alignment is only available in the \n2|Compare Pro version. Visit www.hopie.tech for more information.");
                msgBox.ShowDialog();
            }
            else
            {

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
                        });
                    }
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
                inForceAlignMode = false;
                //DocCompareLeftStatsGrid.Visibility = Visibility.Collapsed;
                //DocCompareRightStatsGrid.Visibility = Visibility.Collapsed;

                SelectDocToCompare win = new SelectDocToCompare();
                List<string> filenames = new List<string>();
                for (int i = 0; i < docs.documents.Count; i++)
                {
                    if (i != docs.documentsToShow[0])
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

                UpdateDocCompareComboBox();
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                docs.forceAlignmentIndices = new List<List<int>>();
                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();

            }
        }

        private void SidePanelOpenDocButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisiblePanel(SidePanels.DRAGDROP);
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

                if (ok == true)
                    items.Add(Path.GetFileName(docs.documents[i].filePath));

                if (i == docs.documentsToShow[0])
                {
                    ind = items.Count - 1;
                }
            }
            Doc1NameLabelComboBox.ItemsSource = items;
            Doc1NameLabelComboBox.SelectedIndex = ind;
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

        }

        private void Doc1NameLabelComboBox_DropDownOpened(object sender, EventArgs e)
        {

        }

        private void Doc1NameLabelComboBox_DropDownClosed(object sender, EventArgs e)
        {

        }

        private void OpenDoc2OriginalButton_Click_1(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + docs.documents[docs.documentsToCompare[1]].filePath + "\"";
            fileopener.Start();
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
            HideDragDropZone3();
            HideDragDropZone4();
            HideDragDropZone5();
            Doc1Grid.Visibility = Visibility.Hidden;
            Doc2Grid.Visibility = Visibility.Hidden;
            Doc3Grid.Visibility = Visibility.Hidden;
            Doc4Grid.Visibility = Visibility.Hidden;
            Doc5Grid.Visibility = Visibility.Hidden;
            //Doc1PageNumberLabel.Content = "";
            Doc2NameLabel.Content = "";
            Doc3NameLabel.Content = "";
            Doc4NameLabel.Content = "";
            Doc5NameLabel.Content = "";

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

            string name = (sender as Button).Name;
            string[] splitName = name.Split("Button");

            switch (int.Parse(splitName[^1]))
            {
                case 1:
                    linkScroll[0] = true;
                    UnlinkScrollButton1.Visibility = Visibility.Visible;
                    LinkScrollButton1.Visibility = Visibility.Hidden;
                    break;
                case 2:
                    linkScroll[1] = true;
                    UnlinkScrollButton2.Visibility = Visibility.Visible;
                    LinkScrollButton2.Visibility = Visibility.Hidden;
                    break;
                case 3:
                    linkScroll[2] = true;
                    UnlinkScrollButton3.Visibility = Visibility.Visible;
                    LinkScrollButton3.Visibility = Visibility.Hidden;
                    break;
                case 4:
                    linkScroll[3] = true;
                    UnlinkScrollButton4.Visibility = Visibility.Visible;
                    LinkScrollButton4.Visibility = Visibility.Hidden;
                    break;
                case 5:
                    linkScroll[4] = true;
                    UnlinkScrollButton5.Visibility = Visibility.Visible;
                    LinkScrollButton5.Visibility = Visibility.Hidden;
                    break;
            }

            int count = 0;
            for (int i = 0; i < docs.documents.Count; i++)
            {
                if (linkScroll[i] == true)
                    count++;
            }

            if (count == 1) // if only one non ref is linked
            {
                if (linkScroll[0] == false) // set ref to linked
                {
                    linkScroll[0] = true;
                    UnlinkScrollButton1.Visibility = Visibility.Visible;
                    LinkScrollButton1.Visibility = Visibility.Hidden;
                }
                else // we make the next document linked?
                {
                    //if(settings.maxDocCount == 2)
                    {
                        linkScroll[1] = true;
                        UnlinkScrollButton2.Visibility = Visibility.Visible;
                        LinkScrollButton2.Visibility = Visibility.Hidden;
                    }
                }
            }

            // trigger a scroll
            Border border = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
            ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

            for (int i = 1; i < docs.documents.Count; i++)
            {
                if (linkScroll[i] == true)
                {
                    Border border2;
                    if (i == 1)
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                    else if (i == 2)
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                    else if (i == 3)
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                    else
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);

                    ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                    //if (scrollViewer.ScrollableHeight > scrollViewer2.ScrollableHeight)
                    //{
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.VerticalOffset + 1);
                        scrollViewer2.ScrollToVerticalOffset(scrollViewer2.VerticalOffset - 1);
                    //}
                    //else
                    //{
                        scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset + 1);
                        scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - 1);
                    //}
                }
            }
        }

        private void UnlinkScrollButton_Click(object sender, RoutedEventArgs e)
        {
            string name = (sender as Button).Name;
            string[] splitName = name.Split("Button");

            switch (int.Parse(splitName[^1]))
            {
                case 1:
                    linkScroll[0] = false;
                    UnlinkScrollButton1.Visibility = Visibility.Hidden;
                    LinkScrollButton1.Visibility = Visibility.Visible;
                    break;
                case 2:
                    linkScroll[1] = false;
                    UnlinkScrollButton2.Visibility = Visibility.Hidden;
                    LinkScrollButton2.Visibility = Visibility.Visible;
                    break;
                case 3:
                    linkScroll[2] = false;
                    UnlinkScrollButton3.Visibility = Visibility.Hidden;
                    LinkScrollButton3.Visibility = Visibility.Visible;
                    break;
                case 4:
                    linkScroll[3] = false;
                    UnlinkScrollButton4.Visibility = Visibility.Hidden;
                    LinkScrollButton4.Visibility = Visibility.Visible;
                    break;
                case 5:
                    linkScroll[4] = false;
                    UnlinkScrollButton5.Visibility = Visibility.Hidden;
                    LinkScrollButton5.Visibility = Visibility.Visible;
                    break;
            }

            int count = 0;
            for(int i= 0; i < docs.documents.Count; i++)
            {
                if (linkScroll[i] == true)
                    count++;
            }

            if(count == 1) // set all to unlink
            {
                for(int i = 0; i < docs.documents.Count; i++)
                {
                    if(linkScroll[i] == true)
                    {
                        switch (i)
                        {
                            case 0:
                                linkScroll[0] = false;
                                UnlinkScrollButton1.Visibility = Visibility.Hidden;
                                LinkScrollButton1.Visibility = Visibility.Visible;
                                break;
                            case 1:
                                linkScroll[1] = false;
                                UnlinkScrollButton2.Visibility = Visibility.Hidden;
                                LinkScrollButton2.Visibility = Visibility.Visible;
                                break;
                            case 2:
                                linkScroll[2] = false;
                                UnlinkScrollButton3.Visibility = Visibility.Hidden;
                                LinkScrollButton3.Visibility = Visibility.Visible;
                                break;
                            case 3:
                                linkScroll[3] = false;
                                UnlinkScrollButton4.Visibility = Visibility.Hidden;
                                LinkScrollButton4.Visibility = Visibility.Visible;
                                break;
                            case 4:
                                linkScroll[4] = false;
                                UnlinkScrollButton5.Visibility = Visibility.Hidden;
                                LinkScrollButton5.Visibility = Visibility.Visible;
                                break;
                        }

                    }
                }
            }
        }

        private void HandleDocPreviewMouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                    lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE &&
                    lic.GetLicenseStatus() != LicenseManagement.LicenseStatus.INACTIVE)
                {
                    Grid item = sender as Grid;
                    Border child = item.Children[0] as Border;
                    Grid childGrid = child.Child as Grid;
                    Image img = childGrid.Children[0] as Image;
                    hiddenPPTEffect = img.Effect;
                    img.Effect = null;

                    hiddenPPTVisi = (childGrid.Children[5] as Label).Visibility;
                    System.Windows.Shapes.Path path = childGrid.Children[4] as System.Windows.Shapes.Path;
                    //path.Visibility = Visibility.Hidden;
                    Label label = childGrid.Children[5] as Label;
                    //label.Visibility = Visibility.Hidden;
                    Grid grid = childGrid.Children[2] as Grid;
                    grid.Visibility = Visibility.Hidden;
                }
                else
                {
                    /*
                    Grid item = sender as Grid;
                    if ((item.Children[3] as System.Windows.Shapes.Path).Visibility == Visibility.Visible)
                    {

                        if (!floatingTip.IsOpen) { floatingTip.IsOpen = true; }

                        Point currentPos = e.GetPosition(outerBorder);

                        // The + 20 part is so your mouse pointer doesn't overlap.
                        floatingTip.HorizontalOffset = currentPos.X + 20;
                        floatingTip.VerticalOffset = currentPos.Y;
                    }
                    */
                }

            }
            catch
            {

            }
        }

        private void HandleDocPreviewMouseLeave(object sender, MouseEventArgs e)
        {
            try
            {
                if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                    lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE)
                {
                    Grid item = sender as Grid;
                    Border child = item.Children[0] as Border;
                    Grid childGrid = child.Child as Grid;
                    Image img = childGrid.Children[0] as Image;
                    img.Effect = hiddenPPTEffect;

                    System.Windows.Shapes.Path path = childGrid.Children[4] as System.Windows.Shapes.Path;
                    path.Visibility = hiddenPPTVisi;
                    Label label = childGrid.Children[5] as Label;
                    label.Visibility = hiddenPPTVisi;
                    Grid grid = childGrid.Children[2] as Grid;
                    grid.Visibility = hiddenPPTVisi;
                }
                else
                {
                    //floatingTip.IsOpen = false;
                }

            }
            catch
            {

            }
        }

        private void HandleShowPPTNoteButton(object sender, RoutedEventArgs e)
        {
            try
            {
                if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                    lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE &&
                    lic.GetLicenseStatus() != LicenseManagement.LicenseStatus.INACTIVE)
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

                    (sender as Button).ToolTip = "Click to show speaker notes";

                }
                else
                {
                    (sender as Button).ToolTip = "Viewing of speaker notes is only available in the 2|Compare pro version.";
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("2|Compare Pro", "Viewing of speaker notes is only available in the \n2|Compare Pro version. Visit www.hopie.tech for more information.");
                    msgBox.ShowDialog();
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

                int i = int.Parse(splitedName[^1]);

                if (docs.pptSpeakerNotesDiff.Count != 0)
                {
                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                    {
                        foreach (DocCompareDLL.Diff diff in docs.pptSpeakerNotesDiff[i])
                        {
                            if (diff.operation == DocCompareDLL.Operation.INSERT)
                            {
                                isChanged |= true;
                            }
                            else if (diff.operation == DocCompareDLL.Operation.DELETE)
                            {
                                isChanged |= true;
                            }
                        }

                    }
                }

                if (splitedName != null)
                {
                    CompareMainItem thisItem = (CompareMainItem)DocCompareMainListView.Items[int.Parse(splitedName[^1])];

                    thisItem.PPTNoteGridLeftVisi = Visibility.Hidden;
                    thisItem.PPTNoteGridRightVisi = Visibility.Hidden;
                    bool showSpeakerNotesLeft = false;
                    bool showSpeakerNotesLeftChanged = false;
                    bool showSpeakerNotesRight = false;
                    bool showSpeakerNotesRightChanged = false;
                    bool showSpeakerNotesRightChangedTrans = false;
                    bool didChange = false;

                    if (docs.documents[docs.documentsToCompare[0]].fileType == Document.FileTypes.PPT && docs.documents[docs.documentsToCompare[1]].fileType == Document.FileTypes.PPT)
                    {
                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1 && docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] == -1)
                        {
                            if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                            {
                                thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n<text>" + docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]] + "</text>";
                                showSpeakerNotesLeft = true;
                                showSpeakerNotesLeftChanged = false;
                            }
                            else
                            {
                                thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n" + "<text></text>";
                                showSpeakerNotesLeft = false;
                                showSpeakerNotesLeftChanged = false;
                            }

                        }
                        else if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] == -1 && docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                        {
                            if (docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                            {
                                thisItem.Document2 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n<text>" + docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]] + "</text>";
                                showSpeakerNotesRight = true;
                                showSpeakerNotesRightChanged = false;
                            }
                            else
                            {
                                thisItem.Document2 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n" + "<text></text>";
                                showSpeakerNotesRight = false;
                                showSpeakerNotesRightChanged = false;
                            }
                        }
                        else if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1 && docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                        {
                            string doc;


                            if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0)
                            {
                                thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n<text>" + docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]] + "</text>";
                            }
                            else
                            {
                                thisItem.Document1 = "<?xml version=\"1.0\" encoding=\"UTF-8\"?> \n" + "<text></text>";
                            }

                            if (docs.pptSpeakerNotesDiff[i].Count != 0)
                            {
                                doc = "<?xml version =\"1.0\" encoding=\"UTF-8\"?> \n<text>";

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

                            if (didChange == false)
                            {
                                if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length == 0
                                    && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length == 0)
                                {
                                    showSpeakerNotesLeft = false;
                                    showSpeakerNotesLeftChanged = false;
                                    showSpeakerNotesRight = false;
                                    showSpeakerNotesRightChanged = false;
                                }
                                else
                                {
                                    showSpeakerNotesLeft = true;
                                    showSpeakerNotesLeftChanged = false;
                                    showSpeakerNotesRight = true;
                                    showSpeakerNotesRightChanged = false;
                                }
                            }
                            else
                            {
                                if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0
                                    && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                                {
                                    showSpeakerNotesLeft = true;
                                    showSpeakerNotesLeftChanged = false;
                                    showSpeakerNotesRight = false;
                                    showSpeakerNotesRightChanged = true;
                                }
                                else if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length == 0
                                    && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length == 0)
                                {
                                    showSpeakerNotesLeft = false;
                                    showSpeakerNotesLeftChanged = false;
                                    showSpeakerNotesRight = false;
                                    showSpeakerNotesRightChanged = false;
                                }
                                else if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length == 0
                                    && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length != 0)
                                {
                                    showSpeakerNotesLeft = false;
                                    showSpeakerNotesLeftChanged = false;
                                    showSpeakerNotesRight = false;
                                    showSpeakerNotesRightChanged = true;
                                }
                                else if (docs.documents[docs.documentsToCompare[0]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[0]].docCompareIndices[i]].Length != 0
                                    && docs.documents[docs.documentsToCompare[1]].pptSpeakerNotes[docs.documents[docs.documentsToCompare[1]].docCompareIndices[i]].Length == 0)
                                {
                                    showSpeakerNotesLeft = true;
                                    showSpeakerNotesLeftChanged = false;
                                    showSpeakerNotesRight = false;
                                    showSpeakerNotesRightChanged = false;
                                    showSpeakerNotesRightChangedTrans = true;
                                }
                                else
                                {
                                    showSpeakerNotesLeft = false;
                                    showSpeakerNotesLeftChanged = true;
                                    showSpeakerNotesRight = false;
                                    showSpeakerNotesRightChanged = true;
                                }
                            }

                        }
                    }

                    if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                                lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                                lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                    {
                        if (didChange == true)
                            thisItem.ShowSpeakerNotesTooltip = "The speaker notes was edited. Get the pro version to view the difference.";
                        else
                            thisItem.ShowSpeakerNotesTooltip = "Viewing of speaker notes is only available in the pro version.";
                    }
                    else
                    {
                        if (didChange == true)
                            thisItem.ShowSpeakerNotesTooltip = "The speaker notes was edited. Click to view the difference.";
                        else
                            thisItem.ShowSpeakerNotesTooltip = "Click to show speaker notes";
                    }

                    if (showSpeakerNotesLeft)
                    {
                        thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Visible;
                        thisItem.showPPTSpeakerNotesButtonLeftChanged = Visibility.Hidden;
                    }
                    else if (showSpeakerNotesLeftChanged)
                    {
                        thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonLeftChanged = Visibility.Visible;
                    }
                    else
                    {
                        thisItem.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonLeftChanged = Visibility.Hidden;
                    }

                    if (showSpeakerNotesRight)
                    {
                        thisItem.showPPTSpeakerNotesButtonRight = Visibility.Visible;
                        thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Hidden;
                    }
                    else if (showSpeakerNotesRightChanged)
                    {
                        thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Visible;
                        thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Hidden;
                    }
                    else if (showSpeakerNotesRightChangedTrans)
                    {
                        thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Visible;
                    }
                    else
                    {
                        thisItem.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                        thisItem.showPPTSpeakerNotesButtonRightChangedTrans = Visibility.Hidden;
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
                if (lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.TRIAL &&
                    lic.GetLicenseTypes() != LicenseManagement.LicenseTypes.FREE &&
                    lic.GetLicenseStatus() != LicenseManagement.LicenseStatus.INACTIVE)
                {
                    string senderName = (sender as Button).Tag.ToString();
                    string[] splitedName;
                    bool leftRight = false;

                    if (senderName.Contains("Left"))
                    {
                        splitedName = senderName.Split("Left");
                        try
                        {
                            int i = int.Parse(splitedName[^1]);
                        }
                        catch
                        {
                            splitedName = senderName.Split("LeftChanged");
                        }
                    }
                    else
                    {
                        splitedName = senderName.Split("Right");
                        try
                        {
                            int i = int.Parse(splitedName[^1]);
                        }
                        catch
                        {
                            splitedName = senderName.Split("RightChanged");
                        }

                        leftRight = true;
                    }

                    if (splitedName != null)
                    {
                        CompareMainItem item = (CompareMainItem)DocCompareMainListView.Items[int.Parse(splitedName[^1])];
                        if (leftRight == false)
                        {
                            if (item.PathToImgRight != null)
                                item.PPTNoteGridRightVisi = Visibility.Visible;

                            item.PPTNoteGridLeftVisi = Visibility.Visible;
                            item.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                            item.showPPTSpeakerNotesButtonLeftChanged = Visibility.Hidden;
                            item.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                            item.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                        }
                        else
                        {
                            if (item.PathToImgLeft != null)
                                item.PPTNoteGridLeftVisi = Visibility.Visible;

                            item.showPPTSpeakerNotesButtonLeft = Visibility.Hidden;
                            item.showPPTSpeakerNotesButtonLeftChanged = Visibility.Hidden;
                            item.showPPTSpeakerNotesButtonRight = Visibility.Hidden;
                            item.showPPTSpeakerNotesButtonRightChanged = Visibility.Hidden;
                            item.PPTNoteGridRightVisi = Visibility.Visible;
                        }
                    }

                    (sender as Button).ToolTip = "Click to show speaker notes.";
                }
                else
                {
                    (sender as Button).ToolTip = "Comparing speaker notes is only available in pro version";
                    CustomMessageBox msgBox = new CustomMessageBox();
                    msgBox.Setup("2|Compare Pro", "Viewing of speaker notes is only available in the \n2|Compare Pro version. Visit www.hopie.tech for more information.");
                    msgBox.ShowDialog();
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

            ShowDoc4FileInfoButton.Visibility = Visibility.Hidden;
            HideDoc4FileInfoButton.Visibility = Visibility.Visible;


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

        private void ShowDoc5FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(5);
            ShowDoc5FileInfoButton.Visibility = Visibility.Hidden;
            HideDoc5FileInfoButton.Visibility = Visibility.Visible;

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

                    if (ext == ".lnk")
                    {
                        Shortcut thisShortCut = Shortcut.ReadFromFile(file);
                        string linkedFile = thisShortCut.LinkTargetIDList.Path;
                        ext = Path.GetExtension(linkedFile);

                        if (IsSupportedFile(ext) == false)
                        {
                            ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(linkedFile));
                        }
                        else
                        {
                            if (docs.documents.Find(x => x.filePath == linkedFile) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                            {
                                docs.AddDocument(linkedFile);
                            }
                            else if (docs.documents.Count >= settings.maxDocCount)
                            {
                                ShowMaxDocCountWarningBox();
                                break;
                            }
                            else
                            {
                                ShowExistingDocCountWarningBox(linkedFile);
                            }
                        }
                    }
                    else
                    {
                        if (IsSupportedFile(ext) == false)
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
                Filter = FileFilter,
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
                    if (IsSupportedFile(ext) == false)
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

                if (linkScroll[3] == true)
                {
                    Border border2;
                    if (linkScroll[0] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[1] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[2] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[4] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView5, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

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

                    if (ext == ".lnk")
                    {
                        Shortcut thisShortCut = Shortcut.ReadFromFile(file);
                        string linkedFile = thisShortCut.LinkTargetIDList.Path;
                        ext = Path.GetExtension(linkedFile);

                        if (IsSupportedFile(ext) == false)
                        {
                            ShowInvalidDocTypeWarningBox(ext, Path.GetFileName(linkedFile));
                        }
                        else
                        {
                            if (docs.documents.Find(x => x.filePath == linkedFile) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                            {
                                docs.AddDocument(linkedFile);
                            }
                            else if (docs.documents.Count >= settings.maxDocCount)
                            {
                                ShowMaxDocCountWarningBox();
                                break;
                            }
                            else
                            {
                                ShowExistingDocCountWarningBox(linkedFile);
                            }
                        }
                    }
                    else
                    {
                        if (IsSupportedFile(ext) == false)
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
                Filter = FileFilter,
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
                    if (IsSupportedFile(ext) == false)
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

                if (linkScroll[4] == true)
                {
                    Border border2;
                    if (linkScroll[0] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView1, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[1] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView2, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[2] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView3, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

                        if (scrollViewer.VerticalOffset <= scrollViewer2.ScrollableHeight)
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                        }
                        else
                        {
                            scrollViewer2.ScrollToVerticalOffset(scrollViewer2.ScrollableHeight);
                        }
                    }

                    if (linkScroll[3] == true)
                    {
                        // try to scroll others
                        border2 = (Border)VisualTreeHelper.GetChild(DocCompareListView4, 0);
                        ScrollViewer scrollViewer2 = VisualTreeHelper.GetChild(border2, 0) as ScrollViewer;

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
            if (docCompareRunning == false && docProcessRunning == false)
            {
                CompareDoc2Button.Visibility = Visibility.Visible;
                Border parentBorder = CompareDoc2Button.Parent as Border;
                parentBorder.Visibility = Visibility.Visible;
            }
        }

        private void Doc2Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc2Button.Visibility = Visibility.Hidden;
            Border parentBorder = CompareDoc2Button.Parent as Border;
            parentBorder.Visibility = Visibility.Hidden;
        }

        private void Doc3Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            if (docCompareRunning == false && docProcessRunning == false)
            {
                CompareDoc3Button.Visibility = Visibility.Visible;
                Border parentBorder = CompareDoc3Button.Parent as Border;
                parentBorder.Visibility = Visibility.Visible;
            }
        }

        private void Doc3Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc3Button.Visibility = Visibility.Hidden;
            Border parentBorder = CompareDoc3Button.Parent as Border;
            parentBorder.Visibility = Visibility.Hidden;
        }

        private void Doc4Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            if (docCompareRunning == false && docProcessRunning == false)
            {
                CompareDoc4Button.Visibility = Visibility.Visible;
                Border parentBorder = CompareDoc4Button.Parent as Border;
                parentBorder.Visibility = Visibility.Visible;
            }
        }

        private void Doc4Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc4Button.Visibility = Visibility.Hidden;
            Border parentBorder = CompareDoc4Button.Parent as Border;
            parentBorder.Visibility = Visibility.Hidden;
        }

        private void Doc5Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            if (docCompareRunning == false && docProcessRunning == false)
            {
                CompareDoc5Button.Visibility = Visibility.Visible;
                Border parentBorder = CompareDoc5Button.Parent as Border;
                parentBorder.Visibility = Visibility.Visible;
            }
        }

        private void Doc5Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc5Button.Visibility = Visibility.Hidden;
            Border parentBorder = CompareDoc5Button.Parent as Border;
            parentBorder.Visibility = Visibility.Hidden;
        }

        private void WindowGetProButton_Click(object sender, RoutedEventArgs e)
        {
            if (localetype == "DE")
            {
                ProcessStartInfo info = new ProcessStartInfo("https://de.hopie.tech/store")
                {
                    UseShellExecute = true
                };
                Process.Start(info);
            }
            else
            {
                ProcessStartInfo info = new ProcessStartInfo("https://en.hopie.tech/store")
                {
                    UseShellExecute = true
                };
                Process.Start(info);
            }
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState == WindowState.Normal)
            {
                outerBorder.Margin = new Thickness(0);
            }

            /*
            if(Doc1Grid.ActualWidth < 300)
            {
                ReferenceTopLabel.Visibility = Visibility.Collapsed;
            }
            else
            {
                ReferenceTopLabel.Visibility = Visibility.Visible;
            }
            */
        }

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
            DirectoryInfo di;

            try
            {
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
                #pragma warning disable SYSLIB0006
                //threadCheckUpdate.Abort();
                //threadCheckTrial.Abort();
                #pragma warning restore SYSLIB0006

            }
            catch (Exception ex)
            {

            }

            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            Application.Current.Shutdown();
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

        private void EnableBrowseFile()
        {
            Dispatcher.Invoke(() =>
            {
                BrowseFileButton1.IsEnabled = true;
                BrowseFileButton2.IsEnabled = true;

                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                    lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                   lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                {
                    BrowseFileButton3.IsEnabled = false;
                    BrowseFileButton4.IsEnabled = false;
                    BrowseFileButton5.IsEnabled = false;
                }
                else
                {
                    BrowseFileButton3.IsEnabled = true;
                    BrowseFileButton4.IsEnabled = true;
                    BrowseFileButton5.IsEnabled = true;
                }
            });
        }

        private void DisableBrowseFile()
        {
            Dispatcher.Invoke(() =>
            {
                BrowseFileButton1.IsEnabled = false;
                BrowseFileButton2.IsEnabled = false;
                BrowseFileButton3.IsEnabled = false;
                BrowseFileButton4.IsEnabled = false;
                BrowseFileButton5.IsEnabled = false;
            });
        }


        private void EnableReload()
        {
            Dispatcher.Invoke(() =>
            {
                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                    lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                   lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                {
                    ReloadDoc1Button.IsEnabled = false;
                    ReloadDoc2Button.IsEnabled = false;
                    ReloadDocCompare1Button.IsEnabled = false;
                    ReloadDocCompare2Button.IsEnabled = false;
                    ReloadDoc1Button.ToolTip = "Reloading document is available only in the pro version";
                    ReloadDoc2Button.ToolTip = "Reloading document is available only in the pro version";
                    ReloadDocCompare1Button.ToolTip = "Reloading document is available only in the pro version";
                    ReloadDocCompare2Button.ToolTip = "Reloading document is available only in the pro version";
                }
                else
                {
                    ReloadDoc1Button.IsEnabled = true;
                    ReloadDoc2Button.IsEnabled = true;
                    ReloadDocCompare1Button.IsEnabled = true;
                    ReloadDocCompare2Button.IsEnabled = true;
                    ReloadDoc1Button.ToolTip = "Reload document";
                    ReloadDoc2Button.ToolTip = "Reload document.";
                    ReloadDocCompare1Button.ToolTip = "Reload document.";
                    ReloadDocCompare2Button.ToolTip = "Reload document";
                }

                ReloadDoc3Button.IsEnabled = true;
                ReloadDoc4Button.IsEnabled = true;
                ReloadDoc5Button.IsEnabled = true;
            });

        }

        private void DocPreviewStatGrid_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                Grid element = (Grid)e.Source;
                double total = 0;
                int i;
                for (i = 0; i < element.ColumnDefinitions.Count; i++)
                {
                    if (e.GetPosition(element).X < total)
                    {
                        break;
                    }

                    total += element.ColumnDefinitions[i].ActualWidth;
                }


                if (docs.documents.Count > (i - 1))
                {

                    switch (i - 1)
                    {
                        case 1:
                            CompareDoc2Button.Visibility = Visibility.Visible;
                            (CompareDoc2Button.Parent as Border).Visibility = Visibility.Visible;
                            CompareDoc3Button.Visibility = Visibility.Hidden;
                            (CompareDoc3Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc4Button.Visibility = Visibility.Hidden;
                            (CompareDoc4Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc5Button.Visibility = Visibility.Hidden;
                            (CompareDoc5Button.Parent as Border).Visibility = Visibility.Hidden;
                            break;
                        case 2:
                            CompareDoc2Button.Visibility = Visibility.Hidden;
                            (CompareDoc2Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc3Button.Visibility = Visibility.Visible;
                            (CompareDoc3Button.Parent as Border).Visibility = Visibility.Visible;
                            CompareDoc4Button.Visibility = Visibility.Hidden;
                            (CompareDoc4Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc5Button.Visibility = Visibility.Hidden;
                            (CompareDoc5Button.Parent as Border).Visibility = Visibility.Hidden;
                            break;
                        case 3:
                            CompareDoc2Button.Visibility = Visibility.Hidden;
                            (CompareDoc2Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc3Button.Visibility = Visibility.Hidden;
                            (CompareDoc3Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc4Button.Visibility = Visibility.Visible;
                            (CompareDoc4Button.Parent as Border).Visibility = Visibility.Visible;
                            CompareDoc5Button.Visibility = Visibility.Hidden;
                            (CompareDoc5Button.Parent as Border).Visibility = Visibility.Hidden;
                            break;
                        case 4:
                            CompareDoc2Button.Visibility = Visibility.Hidden;
                            (CompareDoc2Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc3Button.Visibility = Visibility.Hidden;
                            (CompareDoc3Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc4Button.Visibility = Visibility.Hidden;
                            (CompareDoc4Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc5Button.Visibility = Visibility.Visible;
                            (CompareDoc5Button.Parent as Border).Visibility = Visibility.Visible;
                            break;
                        default:
                            CompareDoc2Button.Visibility = Visibility.Hidden;
                            (CompareDoc2Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc3Button.Visibility = Visibility.Hidden;
                            (CompareDoc3Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc4Button.Visibility = Visibility.Hidden;
                            (CompareDoc4Button.Parent as Border).Visibility = Visibility.Hidden;
                            CompareDoc5Button.Visibility = Visibility.Hidden;
                            (CompareDoc5Button.Parent as Border).Visibility = Visibility.Hidden;
                            break;

                    }
                }
                else
                {
                    CompareDoc2Button.Visibility = Visibility.Hidden;
                    (CompareDoc2Button.Parent as Border).Visibility = Visibility.Hidden;
                    CompareDoc3Button.Visibility = Visibility.Hidden;
                    (CompareDoc3Button.Parent as Border).Visibility = Visibility.Hidden;
                    CompareDoc4Button.Visibility = Visibility.Hidden;
                    (CompareDoc4Button.Parent as Border).Visibility = Visibility.Hidden;
                    CompareDoc5Button.Visibility = Visibility.Hidden;
                    (CompareDoc5Button.Parent as Border).Visibility = Visibility.Hidden;
                }
            }
            catch
            {

            }
        }

        private void DocPreviewStatGrid_MouseLeave(object sender, MouseEventArgs e)
        {
            CompareDoc2Button.Visibility = Visibility.Hidden;
            (CompareDoc2Button.Parent as Border).Visibility = Visibility.Hidden;
            CompareDoc3Button.Visibility = Visibility.Hidden;
            (CompareDoc3Button.Parent as Border).Visibility = Visibility.Hidden;
            CompareDoc4Button.Visibility = Visibility.Hidden;
            (CompareDoc4Button.Parent as Border).Visibility = Visibility.Hidden;
            CompareDoc5Button.Visibility = Visibility.Hidden;
            (CompareDoc5Button.Parent as Border).Visibility = Visibility.Hidden;
        }

        private void HideDoc1FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            ShowDoc1FileInfoButton.Visibility = Visibility.Visible;
            HideDoc1FileInfoButton.Visibility = Visibility.Hidden;

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

        private void HideDoc2FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            ShowDoc2FileInfoButton.Visibility = Visibility.Visible;
            HideDoc2FileInfoButton.Visibility = Visibility.Hidden;

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

        private void HideDoc3FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            ShowDoc3FileInfoButton.Visibility = Visibility.Visible;
            HideDoc3FileInfoButton.Visibility = Visibility.Hidden;

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

        private void HideDoc4FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            ShowDoc4FileInfoButton.Visibility = Visibility.Visible;
            HideDoc4FileInfoButton.Visibility = Visibility.Hidden;

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

        private void HideDoc5FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            ShowDoc5FileInfoButton.Visibility = Visibility.Visible;
            HideDoc5FileInfoButton.Visibility = Visibility.Hidden;

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

        private void HideDocCompareFileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            ShowDocCompareFileInfoButton.Visibility = Visibility.Visible;
            HideDocCompareFileInfoButton.Visibility = Visibility.Hidden;

            DocCompareLeftStatAuthorLabel.Visibility = Visibility.Collapsed;
            DocCompareLeftStatCreatedLabel.Visibility = Visibility.Collapsed;
            DocCompareLeftStatAuthorLabel0.Visibility = Visibility.Collapsed;
            DocCompareLeftStatCreatedLabel0.Visibility = Visibility.Collapsed;
            DocCompareLeftStatLastEditorLabel0.Visibility = Visibility.Collapsed;
            DocCompareLeftStatLastEditorLabel.Visibility = Visibility.Collapsed;

            DocCompareRightStatAuthorLabel.Visibility = Visibility.Collapsed;
            DocCompareRightStatCreatedLabel.Visibility = Visibility.Collapsed;
            DocCompareRightStatCreatedLabel0.Visibility = Visibility.Collapsed;
            DocCompareRightStatAuthorLabel0.Visibility = Visibility.Collapsed;
            DocCompareRightStatAuthorLabel0.Visibility = Visibility.Collapsed;
            DocCompareRightStatLastEditorLabel0.Visibility = Visibility.Collapsed;
            DocCompareRightStatLastEditorLabel.Visibility = Visibility.Collapsed;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (currentVisiblePanel != SidePanels.SETTINGS)
            {
                if (e.Key == Key.LeftCtrl)
                {
                    magnifier.Visibility = Visibility.Visible;
                    magnifier.Freeze(false);
                    //magnifier.ZoomFactor = 0.5;
                }
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftCtrl)
            {
                //magnifier.ZoomFactor = 1.0;
                magnifier.Freeze(true);
                magnifier.Visibility = Visibility.Hidden;

                if (enableZoom == true)
                {
                    enableZoom = false;
                    EnableZoomButton.Visibility = Visibility.Visible;
                    DisableZoomButton.Visibility = Visibility.Hidden;
                    ZoomButtonBackground2.Visibility = Visibility.Hidden;
                }
            }
        }

        private void EnableZoomButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentVisiblePanel != SidePanels.SETTINGS)
            {
                enableZoom = true;
                EnableZoomButton.Visibility = Visibility.Hidden;
                DisableZoomButton.Visibility = Visibility.Visible;
                ZoomButtonBackground2.Visibility = Visibility.Visible;
                magnifier.Visibility = Visibility.Visible;
                magnifier.Freeze(false);
                magnifier.ZoomFactor = 0.5;
            }
        }

        private void DisableZoomButton_Click(object sender, RoutedEventArgs e)
        {
            enableZoom = false;
            EnableZoomButton.Visibility = Visibility.Visible;
            DisableZoomButton.Visibility = Visibility.Hidden;
            ZoomButtonBackground2.Visibility = Visibility.Hidden;
            magnifier.Visibility = Visibility.Hidden;
            magnifier.Freeze(true);
            //magnifier.ZoomFactor = 1.0;
        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if(magnifier.Visibility == Visibility.Visible)
            {
                if(e.Delta > 0)
                {
                    if (magnifier.ZoomFactor > 0.3)
                        magnifier.ZoomFactor = magnifier.ZoomFactor - 0.1;
                    else
                        magnifier.ZoomFactor = 0.3;
                }
                else
                {
                    if (magnifier.ZoomFactor < 0.6)
                        magnifier.ZoomFactor = magnifier.ZoomFactor + 0.1;
                    else
                        magnifier.ZoomFactor = 0.7;
                }

                e.Handled = true;
            }
        }

        private void DisableReload()
        {
            Dispatcher.Invoke(() =>
            {
                ReloadDoc1Button.IsEnabled = false;
                ReloadDoc2Button.IsEnabled = false;
                ReloadDocCompare1Button.IsEnabled = false;
                ReloadDocCompare2Button.IsEnabled = false;

                ReloadDoc3Button.IsEnabled = false;
                ReloadDoc4Button.IsEnabled = false;
                ReloadDoc5Button.IsEnabled = false;
            });
        }

        private void ShowMaskButtonGreen_Click(object sender, RoutedEventArgs e)
        {
            showMask = MaskType.Green;

            foreach (CompareMainItem item in DocCompareMainListView.Items)
            {
                item.ShowMaskMagenta = Visibility.Hidden;
                item.ShowMaskGreen = Visibility.Visible;
            }

            foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
            {
                item.ShowMaskMagenta = Visibility.Hidden;
                item.ShowMaskGreen = Visibility.Visible;
            }

            ShowMaskButtonMagenta.Visibility = Visibility.Hidden;
            ShowMaskButtonGreen.Visibility = Visibility.Hidden;
            HideMaskButton.Visibility = Visibility.Visible;
            HighlightingDisableTip.Visibility = Visibility.Hidden;
        }

        private void EnableOpenOriginal()
        {
            Dispatcher.Invoke(() =>
            {
                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.FREE ||
                lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL ||
                   lic.GetLicenseStatus() == LicenseManagement.LicenseStatus.INACTIVE)
                {
                    OpenDoc1OriginalButton.IsEnabled = false;
                    OpenDoc2OriginalButton.IsEnabled = false;
                    OpenDoc1OriginalButton1.IsEnabled = false;
                    OpenDoc2OriginalButton2.IsEnabled = false;
                    OpenDoc1OriginalButton.ToolTip = "Opening document in external editor is available only in the pro version";
                    OpenDoc2OriginalButton.ToolTip = "Opening document in external editor is available only in the pro version";
                    OpenDoc1OriginalButton1.ToolTip = "Opening document in external editor is available only in the pro version";
                    OpenDoc2OriginalButton2.ToolTip = "Opening document in external editor is available only in the pro version";
                }
                else
                {
                    OpenDoc1OriginalButton.IsEnabled = true;
                    OpenDoc2OriginalButton.IsEnabled = true;
                    OpenDoc1OriginalButton1.IsEnabled = true;
                    OpenDoc2OriginalButton2.IsEnabled = true;
                    OpenDoc1OriginalButton.ToolTip = "Open original";
                    OpenDoc2OriginalButton.ToolTip = "Open original";
                    OpenDoc1OriginalButton1.ToolTip = "Open original";
                    OpenDoc2OriginalButton2.ToolTip = "Open original";
                }

                OpenDoc3OriginalButton3.IsEnabled = true;
                OpenDoc4OriginalButton4.IsEnabled = true;
                OpenDoc5OriginalButton5.IsEnabled = true;
            });
        }

        private void DisableOpenOriginal()
        {
            Dispatcher.Invoke(() =>
            {
                OpenDoc1OriginalButton.IsEnabled = false;
                OpenDoc2OriginalButton.IsEnabled = false;
                OpenDoc1OriginalButton1.IsEnabled = false;
                OpenDoc2OriginalButton2.IsEnabled = false;

                OpenDoc3OriginalButton3.IsEnabled = false;
                OpenDoc4OriginalButton4.IsEnabled = false;
                OpenDoc5OriginalButton5.IsEnabled = false;
            });
        }

        private void EnableCloseDocument()
        {
            Dispatcher.Invoke(() =>
            {
                CloseDoc1Button.IsEnabled = true;
                CloseDoc2Button.IsEnabled = true;
                CloseDoc3Button.IsEnabled = true;
                CloseDoc4Button.IsEnabled = true;
                CloseDoc5Button.IsEnabled = true;
            });
        }
        private void DisableCloseDocument()
        {
            Dispatcher.Invoke(() =>
            {
                CloseDoc1Button.IsEnabled = false;
                CloseDoc2Button.IsEnabled = false;
                CloseDoc3Button.IsEnabled = false;
                CloseDoc4Button.IsEnabled = false;
                CloseDoc5Button.IsEnabled = false;
            });
        }

        private void EnableShowDocInfo()
        {
            Dispatcher.Invoke(() =>
            {
                ShowDoc1FileInfoButton.IsEnabled = true;
                ShowDoc2FileInfoButton.IsEnabled = true;
                ShowDoc3FileInfoButton.IsEnabled = true;
                ShowDoc4FileInfoButton.IsEnabled = true;
                ShowDoc5FileInfoButton.IsEnabled = true;
                ShowDocCompareFileInfoButton.IsEnabled = true;
            });
        }
        private void DisableShowDocInfo()
        {
            Dispatcher.Invoke(() =>
            {
                ShowDoc1FileInfoButton.IsEnabled = false;
                ShowDoc2FileInfoButton.IsEnabled = false;
                ShowDoc3FileInfoButton.IsEnabled = false;
                ShowDoc4FileInfoButton.IsEnabled = false;
                ShowDoc5FileInfoButton.IsEnabled = false;
                ShowDocCompareFileInfoButton.IsEnabled = false;
            });
        }
    }
}