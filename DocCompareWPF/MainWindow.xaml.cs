﻿using DocCompareWPF.Classes;
using Microsoft.Win32;
using ProtoBuf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Diagnostics;
using System.IO;
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
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly DocumentManagement docs;

        private readonly string workingDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare");
        private readonly string appDataDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare");
        private readonly string version = "0.4.1";

        // Stack panel for viewing documents in scrollviewer control in comparison view
        //private VirtualizingStackPanel childPanel1;
        private StackPanel refDocPanel;

        private bool docCompareRunning, docProcessRunning, animateDiffRunning, showMask;

        private int docCompareSideGridShown, docProcessingCounter;

        private Grid gridToAnimate;

        private bool inForceAlignMode;

        private string lastUsedDirectory;

        // License management
        private LicenseManagement lic;

        private string licKeyLastInputString;
        private double scrollPosLeft, scrollPosRight;

        private string selectedSideGridButtonName1 = "";

        private string selectedSideGridButtonName2 = "";

        // App settings
        private AppSettings settings;

        private GridSelection sideGridSelectedLeftOrRight, mainGridSelectedLeftOrRight;

        private Thread threadLoadDocs, threadLoadDocsProgress, threadCompare, threadAnimateDiff, threadDisplayResult, threadCheckTrial, threadRenewLic;

        public MainWindow()
        {
            InitializeComponent();
            showMask = true;
            Directory.CreateDirectory(appDataDir);

            // GUI stuff
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
            /*
            Dispatcher.Invoke(() =>
            {
                if (settings.numPanelsDragDrop == 3)
                    SettingsShowThirdPanelCheckBox.IsChecked = true;
                else
                    SettingsShowThirdPanelCheckBox.IsChecked = false;

                SettingsDefaultFolderTextBox.Content = settings.defaultFolder;

                if (settings.isProVersion == true)
                {
                    SettingsShowThirdPanelCheckBox.IsEnabled = true;
                }
                else
                {
                    SettingsShowThirdPanelCheckBox.IsEnabled = false;
                }
            });
            */

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

                ErrorHandling.ReportStatus("App Launch", "App version: " + version + " successfully launched with license: " + lic.GetLicenseTypesString() + ", expires/renewal on " + lic.GetExpiryDateString() + " on " + lic.GetUUID());
            }
            catch
            {
                lic = new LicenseManagement();
                lic.Init(); // init 14 days trial
                DisplayLicense();
                SaveLicense();

                threadCheckTrial = new Thread(new ThreadStart(CheckTrial));
                threadCheckTrial.Start();

                ErrorHandling.ReportError("New trial license", "on " + lic.GetUUID(), "Expires on " + lic.GetExpiryDateString());
            }

            TimeSpan timeBuffer = lic.GetExpiryDate().Subtract(DateTime.Today);

            // Reminder to subscribe
            if (timeBuffer.TotalDays <= 5 && timeBuffer.TotalDays > 0 && lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
            {
                MessageBox.Show("Your trial license will expire in " + timeBuffer.TotalDays + " day(s). Please consider making a subscription on www.hopietech.com", "Expired lincense", MessageBoxButton.OK);
            }

            // if license expires or needs renewal
            if (timeBuffer.TotalDays <= 0)
            {
                if (lic.GetLicenseTypes() == LicenseManagement.LicenseTypes.TRIAL)
                {
                    MessageBox.Show("Your license has expired. Please consider making a subscription on www.hopietech.com", "Expired lincense", MessageBoxButton.OK);

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

        private async void CheckTrial()
        {
            LicenseManagement.LicServerResponse res = await lic.ActivateTrial();
            if (res == LicenseManagement.LicServerResponse.INVALID)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("A previous trial license has been activate don this machine. Please consider making a subscription on www.hopietech.com", "Expired lincense", MessageBoxButton.OK);

                    BrowseFileButton1.IsEnabled = false;
                    DocCompareFirstDocZone.AllowDrop = false;
                    DocCompareDragDropZone1.AllowDrop = false;
                    DocCompareColorZone1.AllowDrop = false;

                    DisplayLicense();
                });
            }
        }

        private async void RenewLicense()
        {
            LicenseManagement.LicServerResponse res = await lic.RenewLincense();

            Dispatcher.Invoke(() =>
            {
                switch (res)
                {
                    case LicenseManagement.LicServerResponse.UNREACHABLE:
                        TimeSpan bufferTime = lic.GetExpiryWaiveDate().Subtract(DateTime.Today);
                        if (bufferTime.TotalDays >= 0)
                        {
                            MessageBox.Show("License server not reachable. Please check your internet connection or launch the application within " + bufferTime.TotalDays.ToString() + " day(s) with working internet connection.", "License server not reachable", MessageBoxButton.OK);
                        }
                        else
                        {
                            MessageBox.Show("License server not reachable. Your license is no longer valid. Please contact us at support@hopietech.com for support if you have previously renewed the subscription.", "License server not reachable", MessageBoxButton.OK);
                            BrowseFileButton1.IsEnabled = false;
                            DocCompareFirstDocZone.AllowDrop = false;
                            DocCompareDragDropZone1.AllowDrop = false;
                            DocCompareColorZone1.AllowDrop = false;
                        }
                        UserEmailTextBox.IsEnabled = true;
                        LicenseKeyTextBox.IsEnabled = true; // after successful activation, we will prevent further editing
                        ActivateLicenseButton.IsEnabled = true;
                        break;

                    case LicenseManagement.LicServerResponse.KEY_MISMATCH:
                        MessageBox.Show("The provided license key does not match the email address. Please check your inputs.", "Invalid license key", MessageBoxButton.OK);
                        UserEmailTextBox.IsEnabled = true;
                        LicenseKeyTextBox.IsEnabled = true; // after successful activation, we will prevent further editing
                        ActivateLicenseButton.IsEnabled = true;
                        break;

                    case LicenseManagement.LicServerResponse.ACCOUNT_NOT_FOUND:
                        MessageBox.Show("No license was found under the given email address. Please check your inputs.", "License not found", MessageBoxButton.OK);
                        UserEmailTextBox.IsEnabled = true;
                        LicenseKeyTextBox.IsEnabled = true; // after successful activation, we will prevent further editing
                        ActivateLicenseButton.IsEnabled = true;
                        break;

                    case LicenseManagement.LicServerResponse.OKAY:
                        MessageBox.Show("License renewed successfully.", "License renewal", MessageBoxButton.OK);
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

        private async void ActivateLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            LicenseManagement.LicServerResponse res = await lic.ActivateLincense(UserEmailTextBox.Text, LicenseKeyTextBox.Text);

            switch (res)
            {
                case LicenseManagement.LicServerResponse.UNREACHABLE:
                    MessageBox.Show("License server not reachable. Please check your internet connection or try again later.", "License server not reachable", MessageBoxButton.OK);
                    break;

                case LicenseManagement.LicServerResponse.KEY_MISMATCH:
                    MessageBox.Show("The provided license key does not match the email address. Please check your inputs.", "Invalid license key", MessageBoxButton.OK);
                    break;

                case LicenseManagement.LicServerResponse.ACCOUNT_NOT_FOUND:
                    MessageBox.Show("No license was found under the given email address. Please check your inputs.", "License not found", MessageBoxButton.OK);
                    break;

                case LicenseManagement.LicServerResponse.OKAY:
                    MessageBox.Show("License activated successfully.", "License activation", MessageBoxButton.OK);
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
                    MessageBox.Show("License activated failed. Please try again later", "License activation failed", MessageBoxButton.OK);
                    // we do nothing and retain current license info
                    break;
                case LicenseManagement.LicServerResponse.INUSE:
                    MessageBox.Show("License has been activated on another machine. Please contact support@hopietech.com for further assitance", "License in used", MessageBoxButton.OK);
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
        }

        private void BrowseFileButton2_Click(object sender, RoutedEventArgs e)
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
            }
            else
            {
                Doc1Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone1.Visibility = Visibility.Visible;
                Doc1StatsGrid.Visibility = Visibility.Collapsed;
            }

            if (docs.documents.Count >= 2)
            {
                if (docs.documentsToShow[1] != -1)
                    DisplayImageMiddle(docs.documentsToShow[1]);

                if (docs.documentsToShow[1] == -1)
                {
                    Doc2Grid.Visibility = Visibility.Hidden;
                    DocCompareDragDropZone2.Visibility = Visibility.Visible;
                    Doc2StatsGrid.Visibility = Visibility.Collapsed;
                }
            }
            else
            {
                Doc2Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone2.Visibility = Visibility.Visible;
                Doc2StatsGrid.Visibility = Visibility.Collapsed;
            }

            if (docs.documents.Count >= 3 && settings.numPanelsDragDrop == 3)
            {
                if (docs.documentsToShow[2] != -1)
                    DisplayImageRight(docs.documentsToShow[2]);

                if (docs.documentsToShow[2] == -1)
                {
                    HideDragDropZone3();
                    Doc3StatsGrid.Visibility = Visibility.Collapsed;
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
                    Doc3StatsGrid.Visibility = Visibility.Collapsed;
                }
            }

            if (docs.documents.Count == 0)
            {
                Doc1Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone1.Visibility = Visibility.Visible;
                Doc1StatsGrid.Visibility = Visibility.Collapsed;
                HideDragDropZone2();
                HideDragDropZone3();
            }
        }

        private void CompareDocsThread()
        {
            try
            {
                docCompareRunning = true;
                int[,] forceIndices = new int[docs.forceAlignmentIndices.Count, 2];
                for (int i = 0; i < docs.forceAlignmentIndices.Count; i++)
                {
                    forceIndices[i, 0] = docs.forceAlignmentIndices[i][0];
                    forceIndices[i, 1] = docs.forceAlignmentIndices[i][1];
                }

                Document.CompareDocs(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[1]].imageFolder, Path.Join(workingDir, "compare"), out docs.pageCompareIndices, out docs.totalLen, forceIndices);
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

                        if (File.Exists(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"))))
                        {
                            thisItem.PathToMaskImgRight = Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"));
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

                    if (File.Exists(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"))))
                    {
                        rightItem.PathToMask = Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"));
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
                    ActivateLicenseButton.IsEnabled = false;
                    break;

                case LicenseManagement.LicenseTypes.TRIAL:
                    LicenseTypeLabel.Content = "Trial license";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    TimeSpan timeBuffer = lic.GetExpiryDate().Subtract(DateTime.Today);
                    if (timeBuffer.TotalDays >= 0)
                        LicenseExpiryLabel.Content = timeBuffer.TotalDays.ToString() + " days";
                    else
                        LicenseExpiryTypeLabel.Content = "Expired";
                    break;

                case LicenseManagement.LicenseTypes.DEVELOPMENT:
                    LicenseTypeLabel.Content = "Developer license";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    LicenseExpiryLabel.Content = "- days";
                    break;

                default:
                    LicenseTypeLabel.Content = "No license found";
                    LicenseExpiryTypeLabel.Content = "Expires in";
                    LicenseExpiryLabel.Content = "- days";
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
            double accuHeight = 0;

            Border border = (Border)VisualTreeHelper.GetChild(DocCompareMainListView, 0);
            ScrollViewer scrollViewer = VisualTreeHelper.GetChild(border, 0) as ScrollViewer;

            for (int i = 0; i < DocCompareMainListView.Items.Count; i++)
            {
                ListViewItem container = DocCompareMainListView.ItemContainerGenerator.ContainerFromItem(DocCompareMainListView.Items[i]) as ListViewItem;
                accuHeight += container.ActualHeight;

                if (accuHeight > scrollViewer.VerticalOffset + scrollViewer.ActualHeight / 3)
                {
                    DocComparePageNumberLabel.Content = (i + 1).ToString() + " / " + DocCompareMainListView.Items.Count.ToString();
                    docCompareSideGridShown = i;
                    HighlightSideGrid();
                    break;
                }
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

        private void DocCompareScrollViewer2_ScrollChanged(object sender, ScrollChangedEventArgs e)
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

        private void DocCompareScrollViewer3_ScrollChanged(object sender, ScrollChangedEventArgs e)
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

        private void DocCompareSideScrollViewerLeft_ScrollChanged(object sender, ScrollChangedEventArgs e)
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

        private void DocCompareSideScrollViewerRight_ScrollChanged(object sender, ScrollChangedEventArgs e)
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
                item.ShowMask = Visibility.Hidden;
            }

            foreach (SideGridItemRight item in DocCompareSideListViewRight.Items)
            {
                item.ShowMask = Visibility.Hidden;
            }

            ShowMaskButton.Visibility = Visibility.Visible;
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
                            File.Exists(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"))))
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
                    ProcessingDocProgressCard.Visibility = Visibility.Hidden;

                    if (docs.documents[docs.documentsToShow[0]].processed == true)
                    {
                        DisplayImageLeft(docs.documentsToShow[0]);
                        ShowDoc1FileInfoButton.IsEnabled = true;
                        OpenDoc1OriginalButton1.IsEnabled = true;
                    }
                    else
                    {
                        if (docs.documents.Count > 1)
                        {
                            docs.documentsToShow[0] = FindNextDocToShow();
                            DisplayImageLeft(docs.documentsToShow[0]);
                            OpenDoc1OriginalButton1.IsEnabled = true;
                            ShowDoc1FileInfoButton.IsEnabled = true;
                        }
                    }

                    if (docs.documents.Count > 1)
                    {
                        if (docs.documents[docs.documentsToShow[1]].processed == true)
                        {
                            DisplayImageMiddle(docs.documentsToShow[1]);
                            OpenDoc2OriginalButton2.IsEnabled = true;
                            ShowDoc2FileInfoButton.IsEnabled = true;
                        }
                        else
                        {
                            if (docs.documents.Count > 2)
                            {
                                docs.documentsToShow[1] = FindNextDocToShow();
                                DisplayImageMiddle(docs.documentsToShow[1]);
                                OpenDoc2OriginalButton2.IsEnabled = true;
                                ShowDoc2FileInfoButton.IsEnabled = true;
                            }
                        }
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

                    ProgressBarDoc1.Visibility = Visibility.Hidden;
                    ProgressBarDoc2.Visibility = Visibility.Hidden;
                    ProgressBarDoc3.Visibility = Visibility.Hidden;

                    BrowseFileTopButton1.IsEnabled = true;
                    BrowseFileTopButton2.IsEnabled = true;
                    BrowseFileTopButton3.IsEnabled = true;

                    ReloadDoc1Button.IsEnabled = true;
                    ReloadDoc2Button.IsEnabled = true;
                    ReloadDoc3Button.IsEnabled = true;

                    CloseDoc1Button.IsEnabled = true;
                    CloseDoc2Button.IsEnabled = true;
                    CloseDoc3Button.IsEnabled = true;

                    OpenDoc1OriginalButton1.IsEnabled = true;
                    OpenDoc2OriginalButton2.IsEnabled = true;
                    OpenDoc3OriginalButton3.IsEnabled = true;

                    UpdateDocSelectionComboBox();

                    if (docs.documents.Count >= 2)
                        SidePanelDocCompareButton.IsEnabled = true;
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
                        docs.documents[i].ClearFolder();
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
                                MessageBox.Show("There was an error converting " + Path.GetFileName(docs.documents[i].filePath) + ". Please repair document and retry.", "PDF File Corruption", MessageBoxButton.OK);
                            else
                                MessageBox.Show("There was an error converting " + Path.GetFileName(docs.documents[i].filePath) + ". Please repair document and retry.", "Powerpoint File Corruption", MessageBoxButton.OK);
                        }
                        else if (ret == -2)
                        {
                            MessageBox.Show("There was an error converting " + Path.GetFileName(docs.documents[i].filePath) + ". No Microsoft PowerPoint installation found.", "Microsoft PowerPoint not found", MessageBoxButton.OK);
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
        }

        private void ReloadDocThread()
        {
            if (docs.documents[docs.docToReload].ReloadDocument() == 0)
            {
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
                    MessageBox.Show("There was an error converting " + Path.GetFileName(docs.documents[docs.docToReload].filePath) + ". Please repair document and retry.", "PDF File Corruption", MessageBoxButton.OK);
                else
                    MessageBox.Show("There was an error converting " + Path.GetFileName(docs.documents[docs.docToReload].filePath) + ". Please repair document and retry.", "Powerpoint File Corruption", MessageBoxButton.OK);
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
            MessageBox.Show("This document has been loaded: " + docName, "Document exists", MessageBoxButton.OK);
        }

        private void UpdateFileStat(int i)
        {
            if (i == 0 && docs.documents.Count >= 1) // DOC1
            {
                Doc1StatAuthorLabel.Content = docs.documents[docs.documentsToShow[0]].Creator;
                Doc1StatCreatedLabel.Content = docs.documents[docs.documentsToShow[0]].CreatedDate;
                Doc1StatLastEditorLabel.Content = docs.documents[docs.documentsToShow[0]].LastEditor;
                Doc1StatModifiedLabel.Content = docs.documents[docs.documentsToShow[0]].ModifiedDate;
            }

            if (i == 1 && docs.documents.Count >= 2) // DOC2
            {
                Doc2StatAuthorLabel.Content = docs.documents[docs.documentsToShow[1]].Creator;
                Doc2StatCreatedLabel.Content = docs.documents[docs.documentsToShow[1]].CreatedDate;
                Doc2StatLastEditorLabel.Content = docs.documents[docs.documentsToShow[1]].LastEditor;
                Doc2StatModifiedLabel.Content = docs.documents[docs.documentsToShow[1]].ModifiedDate;
            }

            if (i == 2 && docs.documents.Count >= 3) // DOC3
            {
                Doc2StatAuthorLabel.Content = docs.documents[docs.documentsToShow[2]].Creator;
                Doc2StatCreatedLabel.Content = docs.documents[docs.documentsToShow[2]].CreatedDate;
                Doc2StatLastEditorLabel.Content = docs.documents[docs.documentsToShow[2]].LastEditor;
                Doc2StatModifiedLabel.Content = docs.documents[docs.documentsToShow[2]].ModifiedDate;
            }

            if (i == 3) // DOC compare
            {
                DocCompareLeftStatAuthorLabel.Content = docs.documents[docs.documentsToCompare[0]].Creator;
                DocCompareLeftStatCreatedLabel.Content = docs.documents[docs.documentsToCompare[0]].CreatedDate;
                DocCompareLeftStatLastEditorLabel.Content = docs.documents[docs.documentsToCompare[0]].LastEditor;
                DocCompareLeftStatModifiedLabel.Content = docs.documents[docs.documentsToCompare[0]].ModifiedDate;
                DocCompareRightStatAuthorLabel.Content = docs.documents[docs.documentsToCompare[1]].Creator;
                DocCompareRightStatCreatedLabel.Content = docs.documents[docs.documentsToCompare[1]].CreatedDate;
                DocCompareRightStatLastEditorLabel.Content = docs.documents[docs.documentsToCompare[1]].LastEditor;
                DocCompareRightStatModifiedLabel.Content = docs.documents[docs.documentsToCompare[1]].ModifiedDate;
            }
        }

        private void ShowDoc1FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(0);

            if (Doc1StatsGrid.Visibility == Visibility.Collapsed)
                Doc1StatsGrid.Visibility = Visibility.Visible;
            else
                Doc1StatsGrid.Visibility = Visibility.Collapsed;
        }

        private void ShowDoc2FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(1);

            if (Doc2StatsGrid.Visibility == Visibility.Collapsed)
                Doc2StatsGrid.Visibility = Visibility.Visible;
            else
                Doc2StatsGrid.Visibility = Visibility.Collapsed;
        }

        private void ShowDoc3FileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(2);

            if (Doc3StatsGrid.Visibility == Visibility.Collapsed)
                Doc3StatsGrid.Visibility = Visibility.Visible;
            else
                Doc3StatsGrid.Visibility = Visibility.Collapsed;
        }

        private void ShowDocCompareFileInfoButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateFileStat(3);

            if (DocCompareLeftStatsGrid.Visibility == Visibility.Collapsed)
                DocCompareLeftStatsGrid.Visibility = Visibility.Visible;
            else
                DocCompareLeftStatsGrid.Visibility = Visibility.Collapsed;

            if (DocCompareRightStatsGrid.Visibility == Visibility.Collapsed)
                DocCompareRightStatsGrid.Visibility = Visibility.Visible;
            else
                DocCompareRightStatsGrid.Visibility = Visibility.Collapsed;
        }

        private void ShowInvalidDocTypeWarningBox(string fileType, string filename)
        {
            MessageBox.Show("Unsupported file type of " + fileType + " selected with " + filename + ". This document will be ignored.", "Unsupported file type", MessageBoxButton.OK);
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
            MessageBox.Show("You have selected more than " + settings.maxDocCount.ToString() + " documents. Only the first " + settings.maxDocCount.ToString() + " documents are loaded.", "Maximum documents loaded", MessageBoxButton.OK);
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
                inForceAlignMode = false;
                DocCompareLeftStatsGrid.Visibility = Visibility.Collapsed;
                DocCompareRightStatsGrid.Visibility = Visibility.Collapsed;

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

        /*
        private void ReleaseDocPreview()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        */
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
                doc.ClearFolder();
                di = new DirectoryInfo(doc.imageFolder);
                di.Delete();
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

    public class SimpleImageItem : INotifyPropertyChanged
    {
        public string PathToFile { get; set; }
        public Thickness Margin { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class CompareMainItem : INotifyPropertyChanged
    {
        public string ImgGridName { get; set; }
        public string ImgGridLeftName { get; set; }
        public string ImgLeftName { get; set; }
        public string ImgAniLeftName { get; set; }
        public string AnimateDiffLeftButtonName { get; set; }
        public string ImgGridRightName { get; set; }
        public string ImgRightName { get; set; }
        public string ImgAniRightName { get; set; }
        public string ImgMaskRightName { get; set; }
        public string AnimateDiffRightButtonName { get; set; }
        public bool AniDiffButtonEnable { get; set; }

        public string PathToImgLeft { get; set; }
        public string PathToAniImgLeft { get; set; }
        public string PathToImgRight { get; set; }
        public string PathToAniImgRight { get; set; }
        public string PathToMaskImgRight { get; set; }

        public Thickness Margin { get; set; }

        private Visibility _showMask;
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

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SideGridItemLeft : INotifyPropertyChanged
    {
        public string GridName { get; set; }
        public string PageNumberLabel { get; set; }
        public string ImgGridName { get; set; }
        public string ImgName { get; set; }
        public string PathToImg { get; set; }

        public string ImgDummyName { get; set; }
        public string PathToImgDummy { get; set; }
        public string ForceAlignButtonName { get; set; }
        public string ForceAlignInvalidButtonName { get; set; }
        public Thickness Margin { get; set; }

        private Color _color;
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

        private Effect _effect;
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

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }

    public class SideGridItemRight : INotifyPropertyChanged
    {
        public string GridName { get; set; }
        public string RemoveForceAlignButtonName { get; set; }

        public Visibility RemoveForceAlignButtonVisibility { get; set; }

        public bool RemoveForceAlignButtonEnable { get; set; }
        public string ImgGridName { get; set; }
        public string ImgName { get; set; }
        public string ImgMaskName { get; set; }
        public string PathToMask { get; set; }
        public string PathToImg { get; set; }

        public string ImgDummyName { get; set; }
        public string PathToImgDummy { get; set; }
        public string ForceAlignButtonName { get; set; }
        public string ForceAlignInvalidButtonName { get; set; }
        private Visibility _showMask;
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
        private Effect _effect;
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
        public Thickness Margin { get; set; }
        private Color _color;
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

        public event PropertyChangedEventHandler PropertyChanged;

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