using DocCompareWPF.Classes;
using Microsoft.Win32;
using ProtoBuf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
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
        enum SidePanels
        {
            DRAGDROP,
            DOCCOMPARE,
            REFDOC,
            FILE_EXPLORER,
            SETTINGS,
        };

        private readonly DocumentManagement docs;
        private readonly string workingDir = Path.Join(Directory.GetCurrentDirectory(), "temp");
        private bool docCompareRunning, showMask;
        private int docCompareSideGridShown;
        private readonly int MAX_DOC_COUNT = 5;

        // Stack panel for viewing documents in scrollviewer control in comparison view
        StackPanel childPanel1, childPanel2, childPanel3, refDocPanel, docCompareChildPanel1, docCompareChildPanel2;
        Thread threadLoadDocs, threadCompare;

        // App settings
        AppSettings settings;
        string lastUsedDirectory;

        public MainWindow()
        {
            InitializeComponent();
            showMask = true;

            // GUI stuff
            SetVisiblePanel(SidePanels.DRAGDROP);
            SidePanelDocCompareButton.IsEnabled = false;

            HideDragDropZone2();
            HideDragDropZone3();

            try
            {
                LoadSettings();
                lastUsedDirectory = settings.defaultFolder;
            }
            catch
            {
                settings = new AppSettings
                {
                    defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    numPanelsDragDrop = 3
                };

                SaveSettings();
            }

            docs = new DocumentManagement(settings.maxDocCount, workingDir, settings);

            Dispatcher.Invoke(() =>
            {
                if (settings.numPanelsDragDrop == 3)
                    SettingsShowThirdPanelCheckBox.IsChecked = true;
                else
                    SettingsShowThirdPanelCheckBox.IsChecked = false;

                SettingsDefaultFolderTextBox.Content = settings.defaultFolder;

                if(settings.isProVersion == true)
                {
                    SettingsShowThirdPanelCheckBox.IsEnabled = true;
                }
                else
                {
                    SettingsShowThirdPanelCheckBox.IsEnabled = false;
                }

            });
        }

        private void SaveSettings()
        {
            using var file = File.Create("AppSettings.bin");
            Serializer.Serialize(file, settings);
        }

        private void LoadSettings()
        {
            settings = new AppSettings();
            using var file = File.OpenRead("AppSettings.bin");
            settings = Serializer.Deserialize<AppSettings>(file);
        }

        private void ShowDragDropZone2()
        {
            DocCompareSecondDocZone.Visibility = Visibility.Visible;
            DragDropPanel.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Star);
        }

        private void HideDragDropZone2()
        {
            DocCompareSecondDocZone.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[1].Width = new GridLength(0, GridUnitType.Star);
        }

        private void ShowDragDropZone3()
        {
            if (settings.numPanelsDragDrop == 3)
            {
                DocCompareThirdDocZone.Visibility = Visibility.Visible;
                DragDropPanel.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);
            }
        }

        private void HideDragDropZone3()
        {
            DocCompareThirdDocZone.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[2].Width = new GridLength(0, GridUnitType.Star);
        }

        private void WindowCloseButton_Click(object sender, RoutedEventArgs e)
        {
            //TODO: implement handling for query before closing
            Close();
        }

        private void WindowMaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowMaximizeButton.Visibility = Visibility.Hidden;
            WindowRestoreButton.Visibility = Visibility.Visible;
            WindowState = WindowState.Maximized;
            MaxHeight = SystemParameters.MaximizedPrimaryScreenHeight - 7;
        }

        private void WindowMinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void WindowRestoreButton_Click(object sender, RoutedEventArgs e)
        {
            WindowMaximizeButton.Visibility = Visibility.Visible;
            WindowRestoreButton.Visibility = Visibility.Hidden;
            WindowState = WindowState.Normal;
        }

        private void SidePanelDocCompareButton_Click(object sender, RoutedEventArgs e)
        {
            if (docs.documents.Count >= 2 && docCompareRunning == false)
            {
                if (settings.isProVersion == true && settings.canSelectRefDoc == true)
                {
                    SetVisiblePanel(SidePanels.REFDOC);

                    // populate list box
                    ObservableCollection<string> items = new ObservableCollection<string>();
                    foreach (Document doc in docs.documents)
                    {
                        items.Add(Path.GetFileName(doc.filePath));
                    }

                    RefDocListBox.ItemsSource = items;
                    RefDocListBox.SelectedIndex = 0;
                }
                else
                {
                    docs.documentsToCompare[0] = 0; // default using the first document selected
                    UpdateDocCompareComboBox();
                    SetVisiblePanel(SidePanels.DOCCOMPARE);
                }
            }
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

        private void DocCompareDragDropZone1_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];

                foreach (string file in data)
                {
                    if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                    {
                        docs.AddDocument(file);
                    }
                    else
                    {
                        ShowMaxDocCountWarningBox();
                        break;
                    }
                }

                LoadFilesCommonPart();

                threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                threadLoadDocs.Start();
            }
        }
        private void DocCompareDragDropZone2_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];

                if (data.Length > settings.maxDocCount)
                    ShowMaxDocCountWarningBox();

                foreach (string file in data)
                {
                    if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                    {
                        docs.AddDocument(file);
                    }
                    else
                    {
                        ShowMaxDocCountWarningBox();
                        break;
                    }
                }

                LoadFilesCommonPart();

                threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                threadLoadDocs.Start();
            }
        }

        private void DocCompareDragDropZone3_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var data = e.Data.GetData(DataFormats.FileDrop) as string[];

                foreach(string file in data)
                {
                    if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                    {
                        docs.AddDocument(file);
                    }
                    else
                    {
                        ShowMaxDocCountWarningBox();
                        break;
                    }
                }

                LoadFilesCommonPart();

                threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                threadLoadDocs.Start();
            }
        }

        private void ProcessDocThread()
        {
            // Going through documents in stack, check if reloading needed
            for (int i = 0; i < docs.documents.Count; i++)
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
                    }
                }
            }

            Dispatcher.Invoke(() =>
            {
                DisplayImageLeft(docs.documentsToShow[0]);
                Dispatcher.Invoke(() =>
                {
                    OpenDoc1OriginalButton1.IsEnabled = true;
                });

                DisplayImageMiddle(docs.documentsToShow[1]);
                Dispatcher.Invoke(() =>
                {
                    OpenDoc2OriginalButton2.IsEnabled = true;
                });

                if (settings.numPanelsDragDrop == 3)
                {
                    if (docs.documents.Count >= 3)
                    {
                        DisplayImageRight(docs.documentsToShow[2]);
                        Dispatcher.Invoke(() =>
                        {
                            OpenDoc3OriginalButton3.IsEnabled = true;
                        });
                    }
                }

                ProgressBarDoc1.Visibility = Visibility.Hidden;
                ProgressBarDoc2.Visibility = Visibility.Hidden;
                ProgressBarDoc3.Visibility = Visibility.Hidden;
                UpdateDocSelectionComboBox();
            });
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

        private void CloseDoc1Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[0], 0);
            UpdateDocSelectionComboBox();
            CloseDocumentCommonPart();
        }

        private void CloseDoc2Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[1], 1);
            UpdateDocSelectionComboBox();
            CloseDocumentCommonPart();
        }

        private void CloseDoc3Button_Click(object sender, RoutedEventArgs e)
        {
            docs.RemoveDocument(docs.documentsToShow[2], 2);
            UpdateDocSelectionComboBox();
            CloseDocumentCommonPart();
        }

        private void ShowMaskButton_Click(object sender, RoutedEventArgs e)
        {
            showMask = true;
            DisplayComparisonResult();
            ShowMaskButton.Visibility = Visibility.Hidden;
            HideMaskButton.Visibility = Visibility.Visible;
            HighlightingDisableTip.Visibility = Visibility.Hidden;
        }

        private void HideMaskButton_Click(object sender, RoutedEventArgs e)
        {
            showMask = false;
            DisplayComparisonResult();
            ShowMaskButton.Visibility = Visibility.Visible;
            HideMaskButton.Visibility = Visibility.Hidden;
            HighlightingDisableTip.Visibility = Visibility.Visible;
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

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisiblePanel(SidePanels.SETTINGS);
        }

        private void SidePanelOpenDocButton_Click(object sender, RoutedEventArgs e)
        {
            SetVisiblePanel(SidePanels.DRAGDROP);
        }

        private void DocCompareMainScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            double accuHeight = 0;

            for (int i = 0; i < docCompareChildPanel1.Children.Count; i++)
            {
                Size currSize = docCompareChildPanel1.Children[i].DesiredSize;
                accuHeight += currSize.Height;

                if (accuHeight > DocCompareMainScrollViewer.VerticalOffset + DocCompareMainScrollViewer.ActualHeight/3)
                {
                    DocComparePageNumberLabel.Content = (i + 1).ToString() + " / " + docCompareChildPanel1.Children.Count.ToString();
                    docCompareSideGridShown = i;
                    HighlightSideGrid();
                    break;
                }
            }
        }

        private void UpdateDocCompareComboBox()
        {
            // update combo box left
            ObservableCollection<string> items = new ObservableCollection<string>();
            for (int i = 0; i < docs.documents.Count; i++)
            {
                if (i != docs.documentsToCompare[0])
                {
                    items.Add(Path.GetFileName(docs.documents[i].filePath));
                }
            }
            DocCompareNameLabel2ComboBox.ItemsSource = items;
            DocCompareNameLabel2ComboBox.SelectedIndex = 0;
        }

        private void UpdateDocSelectionComboBox()
        {
            // update combo box left
            ObservableCollection<string> items = new ObservableCollection<string>();
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
            }
            Doc1NameLabelComboBox.ItemsSource = items;

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
            }
            Doc2NameLabelComboBox.ItemsSource = items;

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
                }
                Doc3NameLabelComboBox.ItemsSource = items;
            }
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
                /*
                docs.documentsToShow[0] = 0;
                DisplayImageLeft(docs.documentsToShow[0]);
                
                */
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
                /*
                docs.documentsToShow[1] = 1;
                DisplayImageMiddle(docs.documentsToShow[1]);
                
                */
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
                /*
                docs.documentsToShow[2] = 1;
                DisplayImageRight(docs.documentsToShow[2]);
                
                */
            }
        }

        private void LoadFilesCommonPart()
        {
            if (docs.documents.Count >= 2)
                SidePanelDocCompareButton.IsEnabled = true;

            if (settings.numPanelsDragDrop == 3)
                docs.documentsToShow = new List<int>() { 0, 1, 2 };
            else
                docs.documentsToShow = new List<int>() { 0, 1 };

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
        }

        private void ShowMaxDocCountWarningBox()
        {
            MessageBox.Show("You have selected more than " + settings.maxDocCount.ToString() + " documents. Only the first " + settings.maxDocCount.ToString() + " documents are loaded. Subscribe to the Pro-version to view unlimited documents.", "Get Pro-Version", MessageBoxButton.OK);
        }

        private void BrowseFileButton1_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF and PPT files (*.pdf, *.ppt)|*.pdf;*.ppt;*.pptx|PDF files (*.pdf)|*.pdf| PPT files (*.ppt)|*.ppt;*pptx|All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filenames = openFileDialog.FileNames;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);

                foreach (string file in filenames)
                {
                    if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                    {
                        docs.AddDocument(file);
                    }
                    else
                    {
                        ShowMaxDocCountWarningBox();
                        break;
                    }
                }

                LoadFilesCommonPart();

                threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                threadLoadDocs.Start();
            }
        }

        private void BrowseFileButton2_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF and PPT files (*.pdf, *.ppt)|*.pdf;*.ppt;*.pptx|PDF files (*.pdf)|*.pdf| PPT files (*.ppt)|*.ppt;*pptx|All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filenames = openFileDialog.FileNames;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);
                
                foreach (string file in filenames)
                {
                    if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                    {
                        docs.AddDocument(file);
                    }
                    else
                    {
                        ShowMaxDocCountWarningBox();
                        break;
                    }
                }

                LoadFilesCommonPart();

                threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                threadLoadDocs.Start();
            }
        }

        private void BrowseFileButton3_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(lastUsedDirectory) == false)
                lastUsedDirectory = settings.defaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF and PPT files (*.pdf, *.ppt)|*.pdf;*.ppt;*.pptx|PDF files (*.pdf)|*.pdf| PPT files (*.ppt)|*.ppt;*pptx|All files|*.*",
                InitialDirectory = lastUsedDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string[] filenames = openFileDialog.FileNames;
                lastUsedDirectory = Path.GetDirectoryName(filenames[0]);
                
                foreach (string file in filenames)
                {
                    if (docs.documents.Find(x => x.filePath == file) == null && docs.documents.Count < settings.maxDocCount) // doc does not exist
                    {
                        docs.AddDocument(file);
                    }
                    else
                    {
                        ShowMaxDocCountWarningBox();
                        break;
                    }
                }

                LoadFilesCommonPart();

                threadLoadDocs = new Thread(new ThreadStart(ProcessDocThread));
                threadLoadDocs.Start();
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

        private void CompareDocsThread()
        {
            try
            {
                docCompareRunning = true;
                Document.CompareDocs(docs.documents[docs.documentsToCompare[0]].imageFolder, docs.documents[docs.documentsToCompare[1]].imageFolder, Path.Join(workingDir, "compare"), out docs.pageCompareIndices, out docs.totalLen);
                docs.documents[docs.documentsToCompare[0]].docCompareIndices = new List<int>();
                docs.documents[docs.documentsToCompare[1]].docCompareIndices = new List<int>();

                if (docs.totalLen != 0) // ? comparion successful
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

        private void RefDocProceedButton_Click(object sender, RoutedEventArgs e)
        {
            docs.documentsToCompare[0] = RefDocListBox.SelectedIndex;
            UpdateDocCompareComboBox();
            SetVisiblePanel(SidePanels.DOCCOMPARE);
        }

        private void RefDocListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int ind = RefDocListBox.SelectedIndex;
                DisplayRefDoc(ind);

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
                docs.documentsToCompare[1] = docs.documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);

                docCompareGrid.Visibility = Visibility.Hidden;
                docCompareSideGridShown = 0;
                DocCompareMainScrollViewer.ScrollToVerticalOffset(0);
                DocCompareSideScrollViewer.ScrollToVerticalOffset(0);
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

        private void SettingsBrowDefaultFolderButton_Click(object sender, RoutedEventArgs e)
        {
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
                            DocCompareSideScrollViewer.ScrollToVerticalOffset(0);
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
                            DocCompareSideScrollViewer.ScrollToVerticalOffset(0);
                            SetVisiblePanel(SidePanels.DOCCOMPARE);
                            ProgressBarDocCompare.Visibility = Visibility.Visible;
                            threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                            threadCompare.Start();
                            break;
                    }
                });
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

            docs.docToReload = docs.documentsToCompare[0];
            docs.displayToReload = 3;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void ReloadDocCompare2Button_Click(object sender, RoutedEventArgs e)
        {
            ProgressBarDocCompareReload.Visibility = Visibility.Visible;

            docs.docToReload = docs.documentsToCompare[1];
            docs.displayToReload = 4;

            threadLoadDocs = new Thread(new ThreadStart(ReloadDocThread));
            threadLoadDocs.Start();
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if(WindowState == WindowState.Normal)
            {
                WindowMaximizeButton.Visibility = Visibility.Visible;
                WindowRestoreButton.Visibility = Visibility.Hidden;
            }
        }

        private void DisplayImageLeft(int docIndex)
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
                            childPanel1.Children.Add(thisImage);
                            pageCounter++;
                        }

                        DocCompareScrollViewer1.Content = childPanel1;
                        DocCompareScrollViewer1.ScrollToVerticalOffset(0);
                        Doc1Grid.Visibility = Visibility.Visible;
                        ProgressBarDoc1.Visibility = Visibility.Hidden;
                    }
                }
            });
        }
        private void DisplayImageMiddle(int docIndex)
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

                        DocCompareScrollViewer2.Content = childPanel2;
                        Doc2Grid.Visibility = Visibility.Visible;
                        DocCompareScrollViewer2.ScrollToVerticalOffset(0);
                        ProgressBarDoc2.Visibility = Visibility.Hidden;
                    }
                }
            });
        }

        private void DisplayImageRight(int docIndex)
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

                        DocCompareScrollViewer3.Content = childPanel3;
                        Doc3Grid.Visibility = Visibility.Visible;
                        DocCompareScrollViewer3.ScrollToVerticalOffset(0);
                        ProgressBarDoc3.Visibility = Visibility.Hidden;
                    }
                }
            });
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

        private void DisplayComparisonResult()
        {
            int pageCounter = 0;
            Brush brush = FindResource("DocumentBackGroundBrush") as Brush;

            Dispatcher.Invoke(() =>
            {
                DocCompareNameLabel1.Content = Path.GetFileName(docs.documents[docs.documentsToCompare[0]].filePath);

                docCompareChildPanel1 = new StackPanel();
                docCompareChildPanel2 = new StackPanel();
                docCompareChildPanel1.Background = brush;
                docCompareChildPanel1.HorizontalAlignment = HorizontalAlignment.Stretch;
                docCompareChildPanel2.Background = brush;
                docCompareChildPanel2.HorizontalAlignment = HorizontalAlignment.Stretch;
                Image thisImage;
                FileStream stream;
                BitmapImage bitmap;

                for (int i = 0; i < docs.totalLen; i++) // going through all the pages of the longest document
                {
                    Grid thisGrid = new Grid();
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc1
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc2

                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1) // doc 1 has a valid page
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
                        thisImage.Margin = new Thickness(10, 10, 10, 10);
                        thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        thisImage.VerticalAlignment = VerticalAlignment.Center;
                        Grid.SetColumn(thisImage, 0);
                        thisGrid.Children.Add(thisImage);
                    }

                    if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1) // doc 2 has a valid page
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
                        thisImage.Margin = new Thickness(10, 10, 10, 10);

                        thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        thisImage.VerticalAlignment = VerticalAlignment.Center;
                        Grid.SetColumn(thisImage, 1);
                        thisGrid.Children.Add(thisImage);

                        if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1 && showMask == true)
                        {
                            if (File.Exists(Path.Join(workingDir, Path.Join("compare", docs.documents[docs.documentsToCompare[0]].docCompareIndices[i].ToString() + "_" + docs.documents[docs.documentsToCompare[1]].docCompareIndices[i].ToString() + ".png"))))
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
                                thisImage.Margin = new Thickness(10, 10, 10, 10);
                                thisImage.HorizontalAlignment = HorizontalAlignment.Stretch;
                                thisImage.VerticalAlignment = VerticalAlignment.Stretch;

                                Grid.SetColumn(thisImage, 2);
                                thisGrid.Children.Add(thisImage);
                            }
                        }
                    }

                    pageCounter++;
                    docCompareChildPanel1.Children.Add(thisGrid);
                }

                DocCompareMainScrollViewer.Content = docCompareChildPanel1;
                pageCounter = 0;

                // side panel
                for (int i = 0; i < docs.totalLen; i++) // going through all the pages of the longest document
                {
                    /*
                    Grid topGrid = new Grid();
                    topGrid.RowDefinitions.Add(new RowDefinition()); // doc1
                    topGrid.RowDefinitions.Add(new RowDefinition()); // page number
                    */
                    Border thisBorder = new Border
                    {
                        BorderBrush = Brushes.Transparent,
                        BorderThickness = new Thickness(0)
                    };
                    Grid thisGrid = new Grid();
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Auto) }); // page number
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc1
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc2
                    thisGrid.MouseLeftButtonDown += (sen, ev) => { HandleMouseClickOnSideScrollView(sen, ev); };
                    thisGrid.Name = "SideGrid" + i.ToString();
                    thisBorder.Child = thisGrid;
                    //Grid.SetRow(thisGrid, 0);

                    if (docs.documents[docs.documentsToCompare[0]].docCompareIndices[i] != -1) // doc 1 has a valid page
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
                        thisImage.Margin = new Thickness(10, 10, 10, 10);

                        thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        Grid.SetColumn(thisImage, 1);
                        thisGrid.Children.Add(thisImage);
                    }

                    if (docs.documents[docs.documentsToCompare[1]].docCompareIndices[i] != -1) // doc 2 has a valid page
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
                        thisImage.Margin = new Thickness(10, 10, 10, 10);

                        double h = bitmap.PixelHeight;
                        double w = bitmap.PixelWidth;
                        thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                        Grid.SetColumn(thisImage, 2);
                        thisGrid.Children.Add(thisImage);
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
                                    thisImage.Margin = new Thickness(10, 10, 10, 10);

                                    //thisImage.HorizontalAlignment = HorizontalAlignment.Stretch;
                                    //thisImage.VerticalAlignment = VerticalAlignment.Stretch;

                                    //thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
                                    Grid.SetColumn(thisImage, 2);
                                    thisGrid.Children.Add(thisImage);
                                }
                                //thisBorder.BorderBrush = Brushes.Red;
                                thisBorder.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                                //thisBorder.BorderThickness = new Thickness(2);
                            }
                        }

                    }

                    pageCounter++;

                    Label thisLabel = new Label
                    {
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Content = (i + 1).ToString(),
                        MinWidth = 20
                    };
                    Grid.SetColumn(thisLabel, 0);

                    thisGrid.Children.Add(thisLabel);
                    //topGrid.Children.Add(thisGrid);
                    docCompareChildPanel2.Children.Add(thisBorder);
                }
                DocCompareSideScrollViewer.Content = docCompareChildPanel2;

                docCompareGrid.Visibility = Visibility.Visible;
                ProgressBarDocCompare.Visibility = Visibility.Hidden;
            });
        }

        private void HandleMouseClickOnSideScrollView(object sender, MouseButtonEventArgs e)
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

        private void HighlightSideGrid()
        {
            Brush brush = FindResource("SideGridActiveBackground") as Brush;
            double accuHeight = 0;
            double windowsHeight = DocCompareSideScrollViewer.ActualHeight;
            Dispatcher.Invoke(() =>
            {
                for (int i = 0; i < docCompareChildPanel2.Children.Count; i++)
                {
                    Border thisBorder;
                    Grid thisGrid;
                    if (i == docCompareSideGridShown)
                    {
                        thisBorder = docCompareChildPanel2.Children[i] as Border;
                        thisGrid = thisBorder.Child as Grid;
                        thisGrid.Background = brush;

                        Size thisSize = thisGrid.DesiredSize;

                        if (accuHeight - windowsHeight / 2 > 0)
                            DocCompareSideScrollViewer.ScrollToVerticalOffset(accuHeight - windowsHeight / 2);
                        else
                            DocCompareSideScrollViewer.ScrollToVerticalOffset(0);

                        accuHeight += thisSize.Height;
                    }
                    else
                    {
                        thisBorder = docCompareChildPanel2.Children[i] as Border;
                        thisGrid = thisBorder.Child as Grid;
                        thisGrid.Background = Brushes.Transparent;

                        Size thisSize = thisGrid.DesiredSize;
                        accuHeight += thisSize.Height;
                    }


                }
            });
        }
    }
}
