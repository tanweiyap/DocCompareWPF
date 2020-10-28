using DocCompareWPF.Classes;
using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using ProtoBuf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
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

        private readonly string workingDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), ".2compare");

        // Stack panel for viewing documents in scrollviewer control in comparison view
        private StackPanel childPanel1, childPanel2, childPanel3, refDocPanel, docCompareChildPanel1, docCompareChildPanelLeft, docCompareChildPanelRight;

        private bool docCompareRunning, docProcessRunning, animateDiffRunning, showMask;

        private int docCompareSideGridShown, docProcessingCounter, pageToAnimate;

        private bool inForceAlignMode;

        private string lastUsedDirectory;

        // License management
        private LicenseManagement lic;

        private double scrollPosLeft, scrollPosRight;

        private string selectedSideGridButtonName1 = "";

        private string selectedSideGridButtonName2 = "";

        private string licKeyLastInputString;

        // App settings
        private AppSettings settings;

        private GridSelection sideGridSelectedLeftOrRight, mainGridSelectedLeftOrRight;

        private Thread threadLoadDocs, threadLoadDocsProgress, threadCompare, threadAnimateDiff;

        public MainWindow()
        {
            InitializeComponent();
            showMask = true;

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
                lic = new LicenseManagement();
                switch (lic.GetLicenseTypes())
                {
                    case LicenseManagement.LicenseTypes.ANNUAL_SUBSCRIPTION:
                        LicenseTypeLabel.Content = "Annual subscription";
                        LicenseStatusTypeLabel.Content = "Renewal on";
                        LicenseStatusLabel.Content = "01.01.2021";
                        break;

                    case LicenseManagement.LicenseTypes.TRIAL:
                        LicenseTypeLabel.Content = "Trial license";
                        LicenseStatusTypeLabel.Content = "Expires in";
                        LicenseStatusLabel.Content = "14 days";
                        break;

                    case LicenseManagement.LicenseTypes.DEVELOPMENT:
                        LicenseTypeLabel.Content = "Developer license";
                        LicenseStatusTypeLabel.Content = "Expires in";
                        LicenseStatusLabel.Content = "- days";
                        break;
                }

                ErrorHandling.ReportError("App Launch", "Launch on " + lic.GetUUID(), "App successfully launched.");
            }
            catch (Exception ex)
            {
                ErrorHandling.ReportException(ex);
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

        private void ActivateLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            int ret = lic.ActivateLincense(UserEmailTextBox.Text, LicenseKeyTextBox.Text);
        }

        private void AnimateDiffThread()
        {
            bool imageToggler = true;

            while (animateDiffRunning)
            {
                Dispatcher.Invoke(() =>
                {
                    // find Grid to animate
                    Grid parentGrid = new Grid();
                    Grid childGrid = new Grid();
                    string[] splittedName;

                    foreach (object child in docCompareChildPanel1.Children)
                    {
                        if (child is Grid)
                        {
                            Grid localGrid = child as Grid;
                            splittedName = localGrid.Name.Split("Grid");
                            if (int.Parse(splittedName[1]) == pageToAnimate)
                            {
                                parentGrid = localGrid;
                                break; // found parent grid
                            }
                        }
                    }

                    // find child grid
                    foreach (object child in parentGrid.Children)
                    {
                        if (child is Grid)
                        {
                            if (mainGridSelectedLeftOrRight == GridSelection.LEFT)
                            {
                                if ((child as Grid).Name.Contains("Left"))
                                {
                                    childGrid = child as Grid;
                                    break;
                                }
                            }
                            else
                            {
                                if ((child as Grid).Name.Contains("Right"))
                                {
                                    childGrid = child as Grid;
                                    break;
                                }
                            }
                        }
                    }

                    // Turn off Mask and animate
                    foreach (object child in childGrid.Children)
                    {
                        if (child is Image)
                        {
                            Image thisImg = child as Image;
                            if (thisImg.Name.Contains("Ani"))
                            {
                                if (imageToggler == false)
                                    thisImg.Visibility = Visibility.Hidden;
                                else
                                    thisImg.Visibility = Visibility.Visible;
                            }
                            else if (thisImg.Name.Contains("Mask"))
                            {
                                //if (imageToggler == false)
                                //    thisImg.Visibility = Visibility.Visible;
                                //else
                                thisImg.Visibility = Visibility.Hidden;
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
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF |All files|*.*",
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
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG")
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
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF |All files|*.*",
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
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG")
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
                Filter = "PDF, PPT and image files (*.pdf, *.ppt, *jpg, *jpeg, *png, *gif)|*.pdf;*.ppt;*.pptx;*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF|PDF files (*.pdf)|*.pdf|PPT files (*.ppt)|*.ppt;*pptx|Image files|*.jpg;*.jpeg;*.JPG;*.JPEG,*.png;*.PNG;*.gif;*.GIF |All files|*.*",
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
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG")
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
            UpdateDocSelectionComboBox();
            CloseDocumentCommonPart();

            if (docs.documents.Count < 2)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void CloseDoc2Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[1], 1);
            UpdateDocSelectionComboBox();
            CloseDocumentCommonPart();

            if (docs.documents.Count < 2)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void CloseDoc3Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[2], 2);
            UpdateDocSelectionComboBox();
            CloseDocumentCommonPart();

            if (docs.documents.Count < 2)
            {
                SidePanelDocCompareButton.IsEnabled = false;
            }
        }

        private void CloseDocumentCommonPart()
        {
            if (docs.documentsToShow[0] != -1)
                DisplayImageLeft(docs.documentsToShow[0]);
            else
            {
                Doc1Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone1.Visibility = Visibility.Visible;
            }

            if (docs.documents.Count >= 2)
            {
                if (docs.documentsToShow[1] != -1)
                    DisplayImageMiddle(docs.documentsToShow[1]);

                if (docs.documentsToShow[1] == -1)
                {
                    Doc2Grid.Visibility = Visibility.Hidden;
                    DocCompareDragDropZone2.Visibility = Visibility.Visible;
                }
            }
            else
            {
                Doc2Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone2.Visibility = Visibility.Visible;
            }

            if (docs.documents.Count >= 3 && settings.numPanelsDragDrop == 3)
            {
                if (docs.documentsToShow[2] != -1)
                    DisplayImageRight(docs.documentsToShow[2]);

                if (docs.documentsToShow[2] == -1)
                {
                    HideDragDropZone3();
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
                    HideDragDropZone3();
            }

            if (docs.documents.Count == 0)
            {
                Doc1Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone1.Visibility = Visibility.Visible;
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
                    DisplayComparisonResult();
                    HighlightSideGrid();
                    ProgressBarDocCompare.Visibility = Visibility.Hidden;
                });
            }
            catch
            {
                docCompareRunning = false;
            }
        }

        private void DisableRemoveForceAlignButton()
        {
            foreach (object obj in docCompareChildPanelRight.Children)
            {
                //Border thisBorder = obj as Border;
                //Grid thisGrid = thisBorder.Child as Grid;
                Grid thisGrid = obj as Grid;
                foreach (object obj2 in thisGrid.Children)
                {
                    if (obj2 is Button)
                    {
                        Button thisButton = obj2 as Button;
                        thisButton.IsEnabled = false;
                    }
                }
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
            int pageCounter = 0;
            Brush brush = FindResource("DocumentBackGroundBrush") as Brush;

            Dispatcher.Invoke(() =>
            {
                DocCompareNameLabel1.Text = Path.GetFileName(docs.documents[docs.documentsToCompare[0]].filePath);

                docCompareChildPanel1 = new StackPanel
                {
                    Background = brush,
                    HorizontalAlignment = HorizontalAlignment.Stretch
                };
                docCompareChildPanelLeft = new StackPanel
                {
                    Background = brush,
                };
                docCompareChildPanelRight = new StackPanel
                {
                    Background = brush,
                };

                Image thisImage;
                FileStream stream;
                BitmapImage bitmap;

                for (int i = 0; i < docs.totalLen; i++) // going through all the pages of the longest document
                {
                    Grid thisGrid = new Grid
                    {
                        IsHitTestVisible = true,
                        Name = "MainImgGrid" + i.ToString(),
                    };
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) }); // doc1
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) }); // doc2

                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1) // doc 1 has a valid page
                    {
                        Grid mainImageGrid = new Grid()
                        {
                            Name = "MainImgGridLeft" + i.ToString(),
                        };

                        thisImage = new Image()
                        {
                            Name = "MainImgLeft" + i.ToString(),
                        };

                        stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = stream;
                        bitmap.EndInit();
                        stream.Close();
                        thisImage.Source = bitmap;
                        thisImage.Margin = new Thickness(10, 10, 10, 10);
                        thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        thisImage.VerticalAlignment = VerticalAlignment.Center;
                        Grid.SetColumn(mainImageGrid, 0);

                        mainImageGrid.Children.Add(thisImage);
                        thisGrid.Children.Add(mainImageGrid);

                        // Image for animating difference
                        if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                        {
                            thisImage = new Image()
                            {
                                Name = "MainImgAniLeft" + i.ToString(),
                                Visibility = Visibility.Hidden,
                            };

                            stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg"));
                            bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.StreamSource = stream;
                            bitmap.EndInit();
                            stream.Close();
                            thisImage.Source = bitmap;
                            thisImage.Margin = new Thickness(10, 10, 10, 10);
                            thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                            thisImage.VerticalAlignment = VerticalAlignment.Center;
                            mainImageGrid.Children.Add(thisImage);

                            Button animateDiffButton = new Button()
                            {
                                Height = 25,
                                Width = 25,
                                Padding = new Thickness(0, 0, 0, 0),
                                Margin = new Thickness(15, 15, 0, 0),
                                ContentTemplate = (DataTemplate)FindResource("AnimateDiffIcon"),
                                Foreground = Brushes.Black,
                                Background = Brushes.White,
                                Opacity = 1.0,
                                Name = "AnimateDiffLeft" + i.ToString(),
                                HorizontalAlignment = HorizontalAlignment.Left,
                                VerticalAlignment = VerticalAlignment.Center,
                                ToolTip = "Click and hold to animate the difference",
                                Visibility = Visibility.Hidden,
                            };

                            animateDiffButton.PreviewMouseDown += (sen, ev) => HandleMainDocCompareAnimateMouseDown(sen, ev);
                            animateDiffButton.PreviewMouseUp += (sen, ev) => HandleMainDocCompareAnimateMouseRelease(sen, ev);
                            mainImageGrid.Children.Add(animateDiffButton);
                        }

                        mainImageGrid.MouseEnter += (sen, ev) => HandleMainDocCompareGridMouseEnter(sen, ev);
                        mainImageGrid.MouseLeave += (sen, ev) => HandleMainDocCompareGridMouseLeave(sen, ev);
                    }

                    if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1) // doc 2 has a valid page
                    {
                        Grid mainImageGrid = new Grid()
                        {
                            Name = "MainImgGridRight" + i.ToString(),
                        };
                        thisImage = new Image()
                        {
                            Name = "MainImgRight" + i.ToString(),
                        };

                        stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = stream;
                        bitmap.EndInit();
                        stream.Close();
                        thisImage.Source = bitmap;
                        thisImage.Margin = new Thickness(10, 10, 10, 10);

                        thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        thisImage.VerticalAlignment = VerticalAlignment.Center;
                        Grid.SetColumn(mainImageGrid, 1);
                        mainImageGrid.Children.Add(thisImage);
                        thisGrid.Children.Add(mainImageGrid);

                        // Image for animating difference
                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                        {
                            thisImage = new Image()
                            {
                                Name = "MainImgAniRight" + i.ToString(),
                                Visibility = Visibility.Hidden,
                            };

                            stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg"));
                            bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.StreamSource = stream;
                            bitmap.EndInit();
                            stream.Close();
                            thisImage.Source = bitmap;
                            thisImage.Margin = new Thickness(10, 10, 10, 10);
                            thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                            thisImage.VerticalAlignment = VerticalAlignment.Center;
                            mainImageGrid.Children.Add(thisImage);
                        }

                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                        {
                            if (File.Exists(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"))) && showMask == true)
                            {
                                thisImage = new Image()
                                {
                                    Name = "MainMaskImgRight" + i.ToString(),
                                };
                                stream = File.OpenRead(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png")));
                                bitmap = new BitmapImage();
                                bitmap.BeginInit();
                                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                bitmap.StreamSource = stream;
                                bitmap.EndInit();
                                stream.Close();
                                thisImage.Source = bitmap;
                                thisImage.Margin = new Thickness(10, 10, 10, 10);
                                thisImage.HorizontalAlignment = HorizontalAlignment.Stretch;
                                thisImage.VerticalAlignment = VerticalAlignment.Stretch;

                                mainImageGrid.Children.Add(thisImage);
                            }

                            Button animateDiffButton = new Button()
                            {
                                Height = 25,
                                Width = 25,
                                Padding = new Thickness(0, 0, 0, 0),
                                Margin = new Thickness(15, 15, 0, 0),
                                ContentTemplate = (DataTemplate)FindResource("AnimateDiffIcon"),
                                Foreground = Brushes.Black,
                                Background = Brushes.White,
                                Opacity = 1.0,
                                Name = "AnimateDiffRight" + i.ToString(),
                                HorizontalAlignment = HorizontalAlignment.Left,
                                VerticalAlignment = VerticalAlignment.Center,
                                ToolTip = "Click and hold to animate the difference",
                                Visibility = Visibility.Hidden,
                            };

                            animateDiffButton.PreviewMouseDown += (sen, ev) => HandleMainDocCompareAnimateMouseDown(sen, ev);
                            animateDiffButton.PreviewMouseUp += (sen, ev) => HandleMainDocCompareAnimateMouseRelease(sen, ev);
                            mainImageGrid.Children.Add(animateDiffButton);
                        }

                        mainImageGrid.MouseEnter += (sen, ev) => HandleMainDocCompareGridMouseEnter(sen, ev);
                        mainImageGrid.MouseLeave += (sen, ev) => HandleMainDocCompareGridMouseLeave(sen, ev);
                    }

                    pageCounter++;
                    docCompareChildPanel1.Children.Add(thisGrid);
                }

                DocCompareMainScrollViewer.Content = docCompareChildPanel1;
                pageCounter = 0;

                // side panel

                for (int i = 0; i < docs.totalLen; i++) // going through all the pages of the longest document
                {
                    Grid thisLeftGrid = new Grid()
                    {
                        Margin = new Thickness(0),
                    };
                    thisLeftGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(25, GridUnitType.Pixel) }); // page number
                    thisLeftGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(115, GridUnitType.Pixel) }); // doc1
                    Grid thisRightGrid = new Grid()
                    {
                        Margin = new Thickness(0),
                    };

                    thisRightGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(25, GridUnitType.Pixel) }); // forcealign icon
                    thisRightGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(115, GridUnitType.Pixel) }); // doc2

                    thisLeftGrid.MouseLeftButtonDown += (sen, ev) => { HandleMouseClickOnSideScrollView(sen, ev); };
                    thisRightGrid.MouseLeftButtonDown += (sen, ev) => { HandleMouseClickOnSideScrollView(sen, ev); };

                    thisLeftGrid.Name = "LeftSideGrid" + i.ToString();
                    thisRightGrid.Name = "RightSideGrid" + i.ToString();
                    bool displayForceAlignButton = true;

                    // force align button
                    foreach (List<int> ind1 in docs.forceAlignmentIndices)
                    {
                        if (ind1[0] == docs.documents[docs.documentsToCompare[0]].docCompareIndices[i])
                        {
                            Button removeForceAlignButton = new Button()
                            {
                                Height = 20,
                                Width = 20,
                                Padding = new Thickness(0, 0, 0, 0),
                                Margin = new Thickness(0),
                                ContentTemplate = (DataTemplate)FindResource("ForceAlignIcon"),
                                Foreground = Brushes.Black,
                                Background = Brushes.White,
                                Opacity = 1,
                                Name = "RemoveForceAlign" + docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString(),
                                IsHitTestVisible = true,
                                HorizontalAlignment = HorizontalAlignment.Center,
                                VerticalAlignment = VerticalAlignment.Center,
                                ToolTip = "Unlink pages",
                            };
                            Grid.SetColumn(removeForceAlignButton, 0);
                            thisRightGrid.Children.Add(removeForceAlignButton);
                            removeForceAlignButton.Click += (sen, ev) => { RemoveForceAlignClicked(sen, ev); };
                        }
                    }

                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1) // doc 1 has a valid page
                    {
                        Grid imageGrid = new Grid()
                        {
                            Name = "SideImageLeft" + i.ToString(),
                        };
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = stream;
                        bitmap.EndInit();
                        stream.Close();
                        thisImage.Source = bitmap;
                        imageGrid.Margin = new Thickness(10, 10, 10, 10);

                        imageGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        Grid.SetColumn(imageGrid, 1);

                        imageGrid.Children.Add(thisImage);
                        thisLeftGrid.Children.Add(imageGrid);

                        if (displayForceAlignButton)
                        {
                            Button forceAlignButtonLeft = new Button()
                            {
                                Height = 25,
                                Width = 25,
                                Padding = new Thickness(0, 0, 0, 0),
                                ContentTemplate = (DataTemplate)FindResource("ForceAlignIcon"),
                                Foreground = Brushes.Black,
                                Background = Brushes.White,
                                Opacity = 1.0,
                                Visibility = Visibility.Hidden,
                                Name = "SideButtonLeft" + i.ToString(),
                                IsHitTestVisible = true,
                                HorizontalAlignment = HorizontalAlignment.Right,
                                VerticalAlignment = VerticalAlignment.Bottom,
                                ToolTip = "Link pages"
                            };

                            forceAlignButtonLeft.Click += (sen, ev) => { SideGridButtonMouseClick(sen, ev); };

                            imageGrid.Children.Add(forceAlignButtonLeft);

                            forceAlignButtonLeft = new Button()
                            {
                                Height = 25,
                                Width = 25,
                                Padding = new Thickness(0, 0, 0, 0),
                                ContentTemplate = (DataTemplate)FindResource("ForceAlignInvalidIcon"),
                                Foreground = Brushes.Black,
                                Background = Brushes.White,
                                Opacity = 1.0,
                                Visibility = Visibility.Hidden,
                                Name = "SideButtonInvalidLeft" + i.ToString(),
                                IsHitTestVisible = true,
                                HorizontalAlignment = HorizontalAlignment.Right,
                                VerticalAlignment = VerticalAlignment.Bottom,
                                ToolTip = "This link would cross a previously set link. Please remove that link before aligning these pages.",
                            };

                            imageGrid.Children.Add(forceAlignButtonLeft);
                        }
                        imageGrid.MouseEnter += SideGridMouseEnter;
                        imageGrid.MouseLeave += SideGridMouseLeave;

                        if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                        {
                            thisImage = new Image();
                            stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg"));
                            bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.StreamSource = stream;
                            bitmap.EndInit();
                            stream.Close();
                            thisImage.Source = bitmap;
                            imageGrid.Margin = new Thickness(10, 10, 10, 10);

                            imageGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                            Grid.SetColumn(imageGrid, 1);

                            thisImage.Visibility = Visibility.Hidden;

                            imageGrid.Children.Add(thisImage);
                        }
                    }
                    else // doc2 has a valid page, we use it as dummy
                    {
                        Grid imageGrid = new Grid()
                        {
                            Name = "SideImageDummyLeft" + i.ToString(),
                        };
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = stream;
                        bitmap.EndInit();
                        stream.Close();
                        thisImage.Source = bitmap;
                        imageGrid.Margin = new Thickness(10, 10, 10, 10);

                        imageGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        Grid.SetColumn(imageGrid, 1);

                        thisImage.Visibility = Visibility.Hidden;

                        imageGrid.Children.Add(thisImage);
                        thisLeftGrid.Children.Add(imageGrid);
                    }

                    if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1) // doc 2 has a valid page
                    {
                        Grid imageGrid = new Grid()
                        {
                            Name = "SideImageRight" + i.ToString(),
                        };
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[1]].imageFolder, docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = stream;
                        bitmap.EndInit();
                        stream.Close();
                        thisImage.Source = bitmap;
                        imageGrid.Margin = new Thickness(10, 10, 10, 10);

                        imageGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        Grid.SetColumn(imageGrid, 1);

                        imageGrid.Children.Add(thisImage);
                        thisRightGrid.Children.Add(imageGrid);

                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                        {
                            if (File.Exists(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"))))
                            {
                                if (showMask == true)
                                {
                                    thisImage = new Image();
                                    stream = File.OpenRead(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png")));
                                    bitmap = new BitmapImage();
                                    bitmap.BeginInit();
                                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                    bitmap.StreamSource = stream;
                                    bitmap.EndInit();

                                    stream.Close();
                                    thisImage.Source = bitmap;
                                    imageGrid.Margin = new Thickness(10, 10, 10, 10);
                                    imageGrid.Children.Add(thisImage);
                                }
                                thisLeftGrid.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                                thisRightGrid.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                            }
                        }

                        if (displayForceAlignButton)
                        {
                            Button forceAlignButtonRight = new Button()
                            {
                                Height = 25,
                                Width = 25,
                                Padding = new Thickness(0, 0, 0, 0),
                                ContentTemplate = (DataTemplate)FindResource("ForceAlignIcon"),
                                Foreground = Brushes.Black,
                                Background = Brushes.White,
                                Opacity = 1.0,
                                Visibility = Visibility.Hidden,
                                Name = "SideButtonRight" + i.ToString(),
                                HorizontalAlignment = HorizontalAlignment.Left,
                                VerticalAlignment = VerticalAlignment.Bottom,
                                ToolTip = "Link pages"
                            };
                            forceAlignButtonRight.Click += (sen, ev) => { SideGridButtonMouseClick(sen, ev); };

                            imageGrid.Children.Add(forceAlignButtonRight);

                            forceAlignButtonRight = new Button()
                            {
                                Height = 25,
                                Width = 25,
                                Padding = new Thickness(0, 0, 0, 0),
                                ContentTemplate = (DataTemplate)FindResource("ForceAlignInvalidIcon"),
                                Foreground = Brushes.Black,
                                Background = Brushes.White,
                                Opacity = 1.0,
                                Visibility = Visibility.Hidden,
                                Name = "SideButtonInvalidRight" + i.ToString(),
                                HorizontalAlignment = HorizontalAlignment.Left,
                                VerticalAlignment = VerticalAlignment.Bottom,
                                ToolTip = "This link would cross a previously set link. Please remove that link before aligning these pages.",
                            };

                            imageGrid.Children.Add(forceAlignButtonRight);
                        }

                        imageGrid.MouseEnter += SideGridMouseEnter;
                        imageGrid.MouseLeave += SideGridMouseLeave;

                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                        {
                            thisImage = new Image();
                            stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg"));
                            bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.StreamSource = stream;
                            bitmap.EndInit();
                            stream.Close();
                            thisImage.Source = bitmap;
                            imageGrid.Margin = new Thickness(10, 10, 10, 10);

                            imageGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                            Grid.SetColumn(imageGrid, 1);

                            thisImage.Visibility = Visibility.Hidden;

                            imageGrid.Children.Add(thisImage);
                        }
                    }
                    else // we use doc 1 as dummy
                    {
                        Grid imageGrid = new Grid()
                        {
                            Name = "SideImageDummyRight" + i.ToString(),
                        };
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = stream;
                        bitmap.EndInit();
                        stream.Close();
                        thisImage.Source = bitmap;
                        imageGrid.Margin = new Thickness(10, 10, 10, 10);

                        imageGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        Grid.SetColumn(imageGrid, 1);

                        thisImage.Visibility = Visibility.Hidden;

                        imageGrid.Children.Add(thisImage);
                        thisRightGrid.Children.Add(imageGrid);
                    }

                    pageCounter++;

                    Label thisLabel = new Label
                    {
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Content = (i + 1).ToString(),
                        MinWidth = 25,
                        FontSize = 12,
                    };
                    Grid.SetColumn(thisLabel, 0);

                    thisLeftGrid.Children.Add(thisLabel);
                    docCompareChildPanelLeft.Children.Add(thisLeftGrid);
                    docCompareChildPanelRight.Children.Add(thisRightGrid);
                }

                DocCompareSideScrollViewerLeft.Content = docCompareChildPanelLeft;
                DocCompareSideScrollViewerRight.Content = docCompareChildPanelRight;

                docCompareGrid.Visibility = Visibility.Visible;
                ProgressBarDocCompare.Visibility = Visibility.Hidden;
                ProgressBarDocCompareAlign.Visibility = Visibility.Hidden;
            });
        }

        private void DisplayImageLeft(int docIndex)
        {
            if (docIndex != -1)
            {
                Brush brush = FindResource("DocumentBackGroundBrush") as Brush;
                Dispatcher.Invoke(() =>
                {
                    if (docs.documents.Count >= 1)
                    {
                        if (docs.documents[docIndex].filePath != null)
                        {
                            int pageCounter = 0;
                            childPanel1 = new StackPanel
                            {
                                Background = brush,
                                HorizontalAlignment = HorizontalAlignment.Stretch,
                                VerticalAlignment = VerticalAlignment.Stretch,
                            };

                            DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                            FileInfo[] fi = di.GetFiles();

                            if (fi.Length != 0)
                            {
                                for (int i = 0; i < fi.Length; i++)
                                {
                                    Image thisImage = new Image();

                                    var stream = File.OpenRead(Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".jpg"));
                                    var bitmap = new BitmapImage();
                                    bitmap.BeginInit();
                                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                    bitmap.StreamSource = stream;
                                    bitmap.EndInit();
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
                                    childPanel1.Children.Add(thisImage);
                                    pageCounter++;
                                }
                            }
                            else
                            {
                                Grid errGrid = new Grid()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Stretch,
                                    VerticalAlignment = VerticalAlignment.Stretch,
                                    Height = DocCompareScrollViewer1.ActualHeight,
                                };

                                Card errCard = new Card()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Center,
                                    VerticalAlignment = VerticalAlignment.Center,
                                    UniformCornerRadius = 5,
                                    Padding = new Thickness(10),
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
                            }

                            DocCompareScrollViewer1.Content = childPanel1;
                            DocCompareScrollViewer1.ScrollToVerticalOffset(0);
                            Doc1Grid.Visibility = Visibility.Visible;
                            ProgressBarDoc1.Visibility = Visibility.Hidden;
                        }
                    }
                });
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
                            childPanel2 = new StackPanel
                            {
                                Background = brush,
                                HorizontalAlignment = HorizontalAlignment.Stretch
                            };

                            DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                            FileInfo[] fi = di.GetFiles();

                            if (fi.Length != 0)
                            {
                                for (int i = 0; i < fi.Length; i++)
                                {
                                    Image thisImage = new Image();
                                    var stream = File.OpenRead(Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".jpg"));
                                    var bitmap = new BitmapImage();
                                    bitmap.BeginInit();
                                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                    bitmap.StreamSource = stream;
                                    bitmap.EndInit();
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
                                    childPanel2.Children.Add(thisImage);
                                    pageCounter++;
                                }
                            }
                            else
                            {
                                Grid errGrid = new Grid()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Stretch,
                                    VerticalAlignment = VerticalAlignment.Stretch,
                                    Height = DocCompareScrollViewer1.ActualHeight,
                                };

                                Card errCard = new Card()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Center,
                                    VerticalAlignment = VerticalAlignment.Center,
                                    UniformCornerRadius = 5,
                                    Padding = new Thickness(10),
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
                            }

                            DocCompareScrollViewer2.Content = childPanel2;
                            Doc2Grid.Visibility = Visibility.Visible;
                            DocCompareScrollViewer2.ScrollToVerticalOffset(0);
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
                            childPanel3 = new StackPanel
                            {
                                Background = brush,
                                HorizontalAlignment = HorizontalAlignment.Stretch
                            };

                            DirectoryInfo di = new DirectoryInfo(docs.documents[docIndex].imageFolder);
                            FileInfo[] fi = di.GetFiles();

                            if (fi.Length != 0)
                            {
                                for (int i = 0; i < fi.Length; i++)
                                {
                                    Image thisImage = new Image();
                                    var stream = File.OpenRead(Path.Join(docs.documents[docIndex].imageFolder, i.ToString() + ".jpg"));
                                    var bitmap = new BitmapImage();
                                    bitmap.BeginInit();
                                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                    bitmap.StreamSource = stream;
                                    bitmap.EndInit();
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
                                    childPanel3.Children.Add(thisImage);
                                    pageCounter++;
                                }
                            }
                            else
                            {
                                Grid errGrid = new Grid()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Stretch,
                                    VerticalAlignment = VerticalAlignment.Stretch,
                                    Height = DocCompareScrollViewer1.ActualHeight,
                                };

                                Card errCard = new Card()
                                {
                                    HorizontalAlignment = HorizontalAlignment.Center,
                                    VerticalAlignment = VerticalAlignment.Center,
                                    UniformCornerRadius = 5,
                                    Padding = new Thickness(10),
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
                            }

                            DocCompareScrollViewer3.Content = childPanel3;
                            Doc3Grid.Visibility = Visibility.Visible;
                            DocCompareScrollViewer3.ScrollToVerticalOffset(0);
                            ProgressBarDoc3.Visibility = Visibility.Hidden;
                        }
                    }
                });
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
            }
            catch
            {
                Doc1NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
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
            }
            catch
            {
                Doc2NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
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
            }
            catch
            {
                Doc3NameLabelComboBox.SelectedIndex = 0;
                UpdateDocSelectionComboBox();
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
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG")
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
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG")
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
                        && ext != ".jpeg" && ext != ".JPG" && ext != ".JPEG" && ext != ".gif" && ext != ".GIF" && ext != ".png" && ext != ".PNG")
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

            for (int i = 0; i < docCompareChildPanel1.Children.Count; i++)
            {
                Size currSize = docCompareChildPanel1.Children[i].DesiredSize;
                accuHeight += currSize.Height;

                if (accuHeight > DocCompareMainScrollViewer.VerticalOffset + DocCompareMainScrollViewer.ActualHeight / 3)
                {
                    DocComparePageNumberLabel.Content = (i + 1).ToString() + " / " + docCompareChildPanel1.Children.Count.ToString();
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
                DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(0);
                DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ProgressBarDocCompare.Visibility = Visibility.Visible;
                });

                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
            catch
            {
            }
        }

        private void DocCompareScrollViewer1_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            double accuHeight = 0;

            for (int i = 0; i < childPanel1.Children.Count; i++)
            {
                Size currSize = childPanel1.Children[i].DesiredSize;
                accuHeight += currSize.Height;

                if (accuHeight > DocCompareScrollViewer1.VerticalOffset)
                {
                    Doc1PageNumberLabel.Content = (i + 1).ToString() + " / " + childPanel1.Children.Count.ToString();
                    break;
                }
            }
        }

        private void DocCompareScrollViewer2_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            double accuHeight = 0;

            for (int i = 0; i < childPanel2.Children.Count; i++)
            {
                Size currSize = childPanel2.Children[i].DesiredSize;
                accuHeight += currSize.Height;

                if (accuHeight > DocCompareScrollViewer2.VerticalOffset)
                {
                    Doc2PageNumberLabel.Content = (i + 1).ToString() + " / " + childPanel2.Children.Count.ToString();
                    break;
                }
            }
        }

        private void DocCompareScrollViewer3_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            double accuHeight = 0;

            for (int i = 0; i < childPanel3.Children.Count; i++)
            {
                Size currSize = childPanel3.Children[i].DesiredSize;
                accuHeight += currSize.Height;

                if (accuHeight > DocCompareScrollViewer3.VerticalOffset)
                {
                    Doc3PageNumberLabel.Content = (i + 1).ToString() + " / " + childPanel3.Children.Count.ToString();
                    break;
                }
            }
        }

        private void DocCompareSideScrollViewerLeft_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (inForceAlignMode == false)
                DocCompareSideScrollViewerRight.ScrollToVerticalOffset(DocCompareSideScrollViewerLeft.VerticalOffset);
            else
            {
                if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                {
                    DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(scrollPosLeft);
                }
            }
        }

        private void DocCompareSideScrollViewerRight_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (inForceAlignMode == false)
                DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(DocCompareSideScrollViewerRight.VerticalOffset);
            else
            {
                if (sideGridSelectedLeftOrRight == GridSelection.RIGHT)
                {
                    DocCompareSideScrollViewerRight.ScrollToVerticalOffset(scrollPosRight);
                }
            }
        }

        private void EnableRemoveForceAlignButton()
        {
            foreach (object obj in docCompareChildPanelRight.Children)
            {
                //Border thisBorder = obj as Border;
                //Grid thisGrid = thisBorder.Child as Grid;
                Grid thisGrid = obj as Grid;
                foreach (object obj2 in thisGrid.Children)
                {
                    if (obj2 is Button)
                    {
                        Button thisButton = obj2 as Button;
                        thisButton.IsEnabled = true;
                    }
                }
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
                    splittedName = (sender as Button).Name.Split("Left");
                }
                else
                {
                    splittedName = (sender as Button).Name.Split("Right");
                }

                pageToAnimate = int.Parse(splittedName[1]);
                animateDiffRunning = true;

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
                        if (thisImg.Name.Contains("Ani"))
                            thisImg.Visibility = Visibility.Hidden;
                        else if (thisImg.Name.Contains("Mask"))
                        {
                            if (showMask == false)
                            {
                                thisImg.Visibility = Visibility.Hidden;
                            }
                            else
                            {
                                thisImg.Visibility = Visibility.Visible;
                            }
                        }
                        else
                            thisImg.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private void HandleMainDocCompareGridMouseEnter(object sender, MouseEventArgs args)
        {
            if (sender is Grid)
            {
                Grid parentGrid = sender as Grid;

                if (parentGrid.Name.Contains("Left"))
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
                            if (thisButton.Name.Contains("Left"))
                                thisButton.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            if (thisButton.Name.Contains("Right"))
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

                if (parentGrid.Name.Contains("Left"))
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
                            if (thisButton.Name.Contains("Left"))
                                thisButton.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            if (thisButton.Name.Contains("Right"))
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
                Grid grid = sender as Grid;
                string[] splittedName = grid.Name.Split("SideGrid");

                double accuHeight = 0;
                for (int i = 0; i < int.Parse(splittedName[1]); i++)
                {
                    Size currSize = docCompareChildPanel1.Children[i].DesiredSize;
                    accuHeight += currSize.Height;
                }

                DocCompareMainScrollViewer.ScrollToVerticalOffset(accuHeight);

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
            DisplayComparisonResult();

            double currOffset = DocCompareMainScrollViewer.VerticalOffset;
            DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
            DocCompareMainScrollViewer.ScrollToVerticalOffset(currOffset);
            //HighlightSideGrid();
            ShowMaskButton.Visibility = Visibility.Visible;
            HideMaskButton.Visibility = Visibility.Hidden;
            HighlightingDisableTip.Visibility = Visibility.Visible;
        }

        private void HighlightSideGrid()
        {
            try
            {
                Brush brush = FindResource("SideGridActiveBackground") as Brush;
                double accuHeight = 0;
                double windowsHeight = DocCompareSideScrollViewerRight.ActualHeight;
                Dispatcher.Invoke(() =>
                {
                    for (int i = 0; i < docCompareChildPanelRight.Children.Count; i++)
                    {
                        Grid thisGrid;
                        if (i == docCompareSideGridShown)
                        {
                            thisGrid = docCompareChildPanelRight.Children[i] as Grid;
                            thisGrid.Background = brush;

                            thisGrid = docCompareChildPanelLeft.Children[i] as Grid;
                            thisGrid.Background = brush;

                            Size thisSize = thisGrid.DesiredSize;

                            if (accuHeight - windowsHeight / 2 > 0)
                                DocCompareSideScrollViewerRight.ScrollToVerticalOffset(accuHeight - windowsHeight / 2);
                            else
                                DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);

                            accuHeight += thisSize.Height;
                        }
                        else
                        {
                            if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1)
                            {
                                if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1)
                                {
                                    if (File.Exists(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"))))
                                    {
                                        thisGrid = docCompareChildPanelRight.Children[i] as Grid;
                                        thisGrid.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));

                                        thisGrid = docCompareChildPanelLeft.Children[i] as Grid;
                                        thisGrid.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                                        Size thisSize = thisGrid.DesiredSize;
                                        accuHeight += thisSize.Height;
                                    }
                                    else
                                    {
                                        thisGrid = docCompareChildPanelRight.Children[i] as Grid;
                                        thisGrid.Background = Brushes.Transparent;

                                        thisGrid = docCompareChildPanelLeft.Children[i] as Grid;
                                        thisGrid.Background = Brushes.Transparent;
                                        Size thisSize = thisGrid.DesiredSize;
                                        accuHeight += thisSize.Height;
                                    }
                                }
                                else
                                {
                                    thisGrid = docCompareChildPanelRight.Children[i] as Grid;
                                    thisGrid.Background = Brushes.Transparent;

                                    thisGrid = docCompareChildPanelLeft.Children[i] as Grid;
                                    thisGrid.Background = Brushes.Transparent;
                                    Size thisSize = thisGrid.DesiredSize;
                                    accuHeight += thisSize.Height;
                                }
                            }
                            else
                            {
                                thisGrid = docCompareChildPanelRight.Children[i] as Grid;
                                thisGrid.Background = Brushes.Transparent;

                                thisGrid = docCompareChildPanelLeft.Children[i] as Grid;
                                thisGrid.Background = Brushes.Transparent;
                                Size thisSize = thisGrid.DesiredSize;
                                accuHeight += thisSize.Height;
                            }
                        }
                    }
                });
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
            });
        }

        private void LoadSettings()
        {
            settings = new AppSettings();
            using var file = File.OpenRead("AppSettings.bin");
            settings = Serializer.Deserialize<AppSettings>(file);
        }

        private void MaskSideGridInForceAlignMode()
        {
            if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
            {
                foreach (object obj in docCompareChildPanelLeft.Children)
                {
                    //Border thisBorder = obj as Border;
                    //Grid thisGrid = thisBorder.Child as Grid;
                    Grid thisGrid = obj as Grid;
                    foreach (object obj2 in thisGrid.Children)
                    {
                        if (obj2 is Grid)
                        {
                            Grid thisTargetGrid = obj2 as Grid;
                            string[] splittedNameTarget;
                            string[] splittedNameRef;

                            splittedNameTarget = thisTargetGrid.Name.Split("Left");
                            splittedNameRef = selectedSideGridButtonName1.Split("Left");

                            if (thisTargetGrid.Name.Contains("Left") && splittedNameTarget[1] != splittedNameRef[1])
                            {
                                thisTargetGrid.Effect = new BlurEffect()
                                {
                                    Radius = 5,
                                };
                            }
                            else
                            {
                                //thisTargetGrid.Effect = null;
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (object obj in docCompareChildPanelRight.Children)
                {
                    //Border thisBorder = obj as Border;
                    //Grid thisGrid = thisBorder.Child as Grid;
                    Grid thisGrid = obj as Grid;
                    foreach (object obj2 in thisGrid.Children)
                    {
                        if (obj2 is Grid)
                        {
                            Grid thisTargetGrid = obj2 as Grid;
                            string[] splittedNameTarget;
                            string[] splittedNameRef;

                            splittedNameTarget = thisTargetGrid.Name.Split("Right");
                            splittedNameRef = selectedSideGridButtonName1.Split("Right");

                            if (thisTargetGrid.Name.Contains("Right") && splittedNameTarget[1] != splittedNameRef[1])
                            {
                                thisTargetGrid.Effect = new BlurEffect()
                                {
                                    Radius = 5,
                                };
                            }
                            else
                            {
                                //thisTargetGrid.Effect = null;
                            }
                        }
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

                    if (docProcessingCounter >= 2 || docProcessRunning == false)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            if (docs.documents[docs.documentsToShow[0]].processed == true)
                            {
                                DisplayImageLeft(docs.documentsToShow[0]);
                                OpenDoc1OriginalButton1.IsEnabled = true;
                            }
                            else
                            {
                                if (docs.documents.Count > 1)
                                {
                                    docs.documentsToShow[0] = FindNextDocToShow();
                                    DisplayImageLeft(docs.documentsToShow[0]);
                                    OpenDoc1OriginalButton1.IsEnabled = true;
                                }
                            }

                            if (docs.documents.Count > 1)
                            {
                                if (docs.documents[docs.documentsToShow[1]].processed == true)
                                {
                                    DisplayImageMiddle(docs.documentsToShow[1]);
                                    OpenDoc2OriginalButton2.IsEnabled = true;
                                }
                                else
                                {
                                    if (docs.documents.Count > 2)
                                    {
                                        docs.documentsToShow[1] = FindNextDocToShow();
                                        DisplayImageMiddle(docs.documentsToShow[1]);
                                        OpenDoc2OriginalButton2.IsEnabled = true;
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
                            UpdateDocSelectionComboBox();
                        });
                    }

                    Thread.Sleep(10);
                }

                Dispatcher.Invoke(() =>
                {
                    ProcessingDocProgressCard.Visibility = Visibility.Hidden;

                    if (docs.documents[docs.documentsToShow[0]].processed == true)
                    {
                        DisplayImageLeft(docs.documentsToShow[0]);
                        OpenDoc1OriginalButton1.IsEnabled = true;
                    }
                    else
                    {
                        if (docs.documents.Count > 1)
                        {
                            docs.documentsToShow[0] = FindNextDocToShow();
                            DisplayImageLeft(docs.documentsToShow[0]);
                            OpenDoc1OriginalButton1.IsEnabled = true;
                        }
                    }

                    if (docs.documents.Count > 1)
                    {
                        if (docs.documents[docs.documentsToShow[1]].processed == true)
                        {
                            DisplayImageMiddle(docs.documentsToShow[1]);
                            OpenDoc2OriginalButton2.IsEnabled = true;
                        }
                        else
                        {
                            if (docs.documents.Count > 2)
                            {
                                docs.documentsToShow[1] = FindNextDocToShow();
                                DisplayImageMiddle(docs.documentsToShow[1]);
                                OpenDoc2OriginalButton2.IsEnabled = true;
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
                            DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                            DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(0);
                            DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);
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
                            DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                            DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(0);
                            DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);
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
                DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(0);
                DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);
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
                string[] splittedName = button.Name.Split("Align");
                docs.RemoveForceAligmentPairs(int.Parse(splittedName[1]));

                ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(0);
                DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);
                SetVisiblePanel(SidePanels.DOCCOMPARE);
                ProgressBarDocCompare.Visibility = Visibility.Visible;
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
        }

        private void SaveSettings()
        {
            using var file = File.Create("AppSettings.bin");
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
                    SettingsBrowseDocButton.Background = brush;
                    SettingsDocCompareButton.Background = Brushes.Transparent;
                    SettingsSubscriptionButton.Background = Brushes.Transparent;
                    SettingsAboutButton.Background = Brushes.Transparent;
                    break;

                case SettingsPanels.ABOUT:
                    SettingsDocBrowsingPanel.Visibility = Visibility.Hidden;
                    SettingsAboutPanel.Visibility = Visibility.Visible;
                    SettingsSubscriptionPanel.Visibility = Visibility.Hidden;
                    SettingsBrowseDocButton.Background = Brushes.Transparent;
                    SettingsDocCompareButton.Background = Brushes.Transparent;
                    SettingsSubscriptionButton.Background = Brushes.Transparent;
                    SettingsAboutButton.Background = brush;
                    break;

                case SettingsPanels.SUBSCRIPTION:
                    SettingsDocBrowsingPanel.Visibility = Visibility.Hidden;
                    SettingsAboutPanel.Visibility = Visibility.Hidden;
                    SettingsSubscriptionPanel.Visibility = Visibility.Visible;
                    SettingsBrowseDocButton.Background = Brushes.Transparent;
                    SettingsDocCompareButton.Background = Brushes.Transparent;
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

        private void ShowInvalidDocTypeWarningBox(string fileType, string filename)
        {
            MessageBox.Show("Unsupported file type of " + fileType + " selected with " + filename + ". This document will be ignored.", "Unsupported file type", MessageBoxButton.OK);
        }

        private void ShowMaskButton_Click(object sender, RoutedEventArgs e)
        {
            showMask = true;
            DisplayComparisonResult();
            double currOffset = DocCompareMainScrollViewer.VerticalOffset;
            DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
            DocCompareMainScrollViewer.ScrollToVerticalOffset(currOffset);
            //HighlightSideGrid();
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
            Button button = sender as Button;
            if (inForceAlignMode == false)
            {
                scrollPosLeft = DocCompareSideScrollViewerLeft.VerticalOffset;
                scrollPosRight = DocCompareSideScrollViewerRight.VerticalOffset;

                selectedSideGridButtonName1 = button.Name;
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
                selectedSideGridButtonName2 = button.Name;

                if (selectedSideGridButtonName2 == selectedSideGridButtonName1) // click on original page again
                {
                    Dispatcher.Invoke(() =>
                    {
                        UnMaskSideGridFromForceAlignMode();
                        EnableRemoveForceAlignButton();

                        if (sideGridSelectedLeftOrRight == GridSelection.LEFT)
                        {
                            //DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(scrollPosLeft);
                            DocCompareSideScrollViewerRight.ScrollToVerticalOffset(scrollPosLeft);
                        }
                        else
                        {
                            DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(scrollPosRight);
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
                    DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                    DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(0);
                    DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);
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
            string imgName = img.Name;
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
                    if ((child as Button).Name == nameToLook)
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
                                        if ((child2 as Button).Name == nameToLook)
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
                                        if ((child2 as Button).Name == nameToLook)
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
                                                if ((child2 as Button).Name == nameToLook)
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
                                                if ((child2 as Button).Name == nameToLook)
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
            string imgName = img.Name;
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
                    if ((child as Button).Name == nameToLook || (child as Button).Name == nameToLook2)
                    {
                        Button foundButton = child as Button;
                        if (inForceAlignMode == false)
                        {
                            foundButton.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            if (foundButton.Name == selectedSideGridButtonName1)
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

                    UpdateDocCompareComboBox();
                    SetVisiblePanel(SidePanels.DOCCOMPARE);
                    docs.forceAlignmentIndices = new List<List<int>>();
                    ProgressBarDocCompareReload.Visibility = Visibility.Hidden;
                    docCompareGrid.Visibility = Visibility.Hidden;
                    docCompareSideGridShown = 0;
                    DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                    DocCompareSideScrollViewerLeft.ScrollToVerticalOffset(0);
                    DocCompareSideScrollViewerRight.ScrollToVerticalOffset(0);
                    ProgressBarDocCompare.Visibility = Visibility.Visible;
                    threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                    threadCompare.Start();
                }
            }
        }

        private void SidePanelOpenDocButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisiblePanel(SidePanels.DRAGDROP);
        }

        private void UnMaskSideGridFromForceAlignMode()
        {
            foreach (object obj in docCompareChildPanelLeft.Children)
            {
                //Border thisBorder = obj as Border;
                //Grid thisGrid = thisBorder.Child as Grid;
                Grid thisGrid = obj as Grid;
                foreach (object obj2 in thisGrid.Children)
                {
                    if (obj2 is Grid)
                    {
                        Grid thisTargetGrid = obj2 as Grid;
                        thisTargetGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 }; ;
                    }
                }
            }

            foreach (object obj in docCompareChildPanelRight.Children)
            {
                //Border thisBorder = obj as Border;
                //Grid thisGrid = thisBorder.Child as Grid;
                Grid thisGrid = obj as Grid;
                foreach (object obj2 in thisGrid.Children)
                {
                    if (obj2 is Grid)
                    {
                        Grid thisTargetGrid = obj2 as Grid;
                        thisTargetGrid.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 }; ;
                    }
                }
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

            di = new DirectoryInfo(Path.Join(workingDir));
            di.Delete(true);

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
}