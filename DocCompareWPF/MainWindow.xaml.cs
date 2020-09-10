using DocCompareWPF.Classes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
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
            FILE_EXPLORER,
            SETTINGS,
        };

        List<Document> documents;
        string workingDir = Path.Join(Directory.GetCurrentDirectory(), "temp");
        bool docCompareRunning, showMask, docProcessRunning;
        ArrayList pageIndices;
        int totalLen, docCompareSideGridShown;
        int MAX_DOC_COUNT = 3;
        int secondDocToShow = 1;
        int firstDocToShow = 0;

        // Stack panel for viewing documents in scrollviewer control
        StackPanel childPanel1, childPanel2, docCompareChildPanel1, docCompareChildPanel2;
        Thread threadDoc1, threadCompare;

        public MainWindow()
        {
            InitializeComponent();
            documents = new List<Document>();
            showMask = true;

            // Add dummy documents
            for (int i = 0; i < MAX_DOC_COUNT; i++)
            {
                documents.Add(new Document());
                Directory.CreateDirectory(Path.Join(workingDir, "doc" + (i + 1).ToString()));
                documents[i].imageFolder = Path.Join(workingDir, "doc" + (i + 1).ToString());
                documents[i].docID = "doc" + (i + 1).ToString();
                documents[i].clearFolder();
            }

            // create the temporary folders for converted images
            Directory.CreateDirectory(Path.Join(workingDir, "compare"));

            // GUI stuff
            SetVisiblePanel(SidePanels.DRAGDROP);
            //OpenDoc1OriginalButton1.IsEnabled = false;
            //OpenDoc2OriginalButton2.IsEnabled = false;

            HideDragDropZone2();
        }

        private void ShowDragDropZone2()
        {
            DocCompareDragDropZone2.Visibility = Visibility.Visible;
            DragDropPanel.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
            DragDropPanel.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Star);
        }

        private void HideDragDropZone2()
        {
            DocCompareDragDropZone2.Visibility = Visibility.Collapsed;
            DragDropPanel.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
            DragDropPanel.ColumnDefinitions[1].Width = new GridLength(0, GridUnitType.Star);
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
            MaxHeight = SystemParameters.MaximizedPrimaryScreenHeight;
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
            if (documents[firstDocToShow].filePath != null && documents[secondDocToShow].filePath != null && docCompareRunning == false)
            {
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
        }

        private void SetVisiblePanel(SidePanels p_sidePanel)
        {
            Brush brush = FindResource("SidePanelActiveBackground") as Brush;

            switch (p_sidePanel)
            {
                case SidePanels.DRAGDROP:
                    SidePanelOpenDocBackground.Background = brush;
                    SidePanelDocCompareBackground.Background = Brushes.Transparent;
                    SettingsButtonBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Visible;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    SettingsPanel.Visibility = Visibility.Hidden;
                    break;
                case SidePanels.DOCCOMPARE:
                    SidePanelOpenDocBackground.Background = Brushes.Transparent;
                    SidePanelDocCompareBackground.Background = brush;
                    SettingsButtonBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Visible;
                    SettingsPanel.Visibility = Visibility.Hidden;
                    break;
                case SidePanels.SETTINGS:
                    SidePanelOpenDocBackground.Background = Brushes.Transparent;
                    SidePanelDocCompareBackground.Background = Brushes.Transparent;
                    SettingsButtonBackground.Background = brush;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    SettingsPanel.Visibility = Visibility.Visible;
                    break;
                default:
                    SidePanelOpenDocBackground.Background = Brushes.Transparent;
                    SidePanelDocCompareBackground.Background = Brushes.Transparent;
                    SettingsButtonBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    SettingsPanel.Visibility = Visibility.Hidden;
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

        private void DocCompareDragDropZone1_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {

                var data = e.Data.GetData(DataFormats.FileDrop) as string[];

                if (data.Length == 1) // only one file drop
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc1Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc1.Visibility = Visibility.Visible;
                        ShowDragDropZone2();

                        //updateDocSelectionComboBox();
                    });

                    documents[firstDocToShow].filePath = data[0];
                    documents[firstDocToShow].loaded = false;
                    documents[firstDocToShow].processed = false;

                }
                else
                {
                    for (int i = 0; i < Math.Min(MAX_DOC_COUNT, data.Length); i++)
                    {
                        documents[i].filePath = data[i];
                        documents[i].loaded = false;
                        documents[i].processed = false;
                    }

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc1Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc1.Visibility = Visibility.Visible;
                        Doc2Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc2.Visibility = Visibility.Visible;
                        //updateDocSelectionComboBox();
                        ShowDragDropZone2();
                    });
                }

                threadDoc1 = new Thread(new ThreadStart(processDocThread));
                threadDoc1.Start();
            }
        }

        private void processDocThread()
        {
            docProcessRunning = true;
            // Going through documents in stack, check if reloading needed
            for (int i = 0; i < documents.Count; i++)
            {
                if (documents[i].loaded == false && documents[i].filePath != null)
                {
                    documents[i].loaded = true;
                    documents[i].clearFolder();
                    documents[i].detectFileType();

                    int ret = -1;
                    switch (documents[i].fileType)
                    {
                        case Document.FileTypes.PDF:
                            ret = documents[i].readPDF();
                            break;
                        case Document.FileTypes.PPT:
                            ret = documents[i].readPPT();
                            break;
                    }
                }
            }

            Dispatcher.Invoke(() =>
            {
                DisplayImageLeft(firstDocToShow);
                Dispatcher.Invoke(() =>
                {
                    OpenDoc1OriginalButton1.IsEnabled = true;
                });

                DisplayImageRight(secondDocToShow);
                Dispatcher.Invoke(() =>
                {
                    OpenDoc2OriginalButton2.IsEnabled = true;
                });

                ProgressBarDoc1.Visibility = Visibility.Hidden;
                ProgressBarDoc2.Visibility = Visibility.Hidden;
                updateDocSelectionComboBox();
            });
            docProcessRunning = false;
        }

        private void CloseDoc1Button_Click(object sender, RoutedEventArgs e)
        {
            string docID = documents[firstDocToShow].docID;
            documents.RemoveAt(firstDocToShow);
            documents.Add(new Document());
            documents[documents.Count - 1].imageFolder = Path.Join(workingDir, docID);
            documents[documents.Count - 1].docID = docID;

            // update combo box
            updateDocSelectionComboBox();
            firstDocToShow = 0;

            if (documents[firstDocToShow].filePath != null)
            {
                DisplayImageLeft(firstDocToShow);
            }
            else
            {
                OpenDoc1OriginalButton1.IsEnabled = false;
                Doc1Grid.Visibility = Visibility.Hidden;
                if (documents[0].filePath == null)
                {
                    HideDragDropZone2();
                }
            }

            if(documents[secondDocToShow].filePath == null)
            {
                secondDocToShow = 1;
                Doc2Grid.Visibility = Visibility.Hidden;
                DocCompareDragDropZone2.Visibility = Visibility.Visible;
            }
        }

        private void CloseDoc2Button_Click(object sender, RoutedEventArgs e)
        {
            string docID = documents[secondDocToShow].docID;
            documents.RemoveAt(secondDocToShow);
            documents.Add(new Document());
            documents[documents.Count - 1].imageFolder = Path.Join(workingDir, docID);
            documents[documents.Count - 1].docID = docID;

            // update combo box
            updateDocSelectionComboBox();
            secondDocToShow = 1;

            if (documents[secondDocToShow].filePath != null)
            {
                DisplayImageRight(secondDocToShow);
            }
            else
            {
                OpenDoc2OriginalButton2.IsEnabled = false;
                Doc2Grid.Visibility = Visibility.Hidden;
                if (documents[0].filePath == null)
                {
                    HideDragDropZone2();
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

        private void ShowMaskButton_Click(object sender, RoutedEventArgs e)
        {
            showMask = true;
            DisplayComparisonResult();
            ShowMaskButton.Visibility = Visibility.Hidden;
            HideMaskButton.Visibility = Visibility.Visible;
        }

        private void HideMaskButton_Click(object sender, RoutedEventArgs e)
        {
            showMask = false;
            DisplayComparisonResult();
            ShowMaskButton.Visibility = Visibility.Visible;
            HideMaskButton.Visibility = Visibility.Hidden;
        }

        private void OpenDoc1OriginalButton_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + documents[firstDocToShow].filePath + "\"";
            fileopener.Start();
        }

        private void OpenDoc2OriginalButton_Click(object sender, RoutedEventArgs e)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + documents[secondDocToShow].filePath + "\"";
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

                if (accuHeight > DocCompareMainScrollViewer.VerticalOffset)
                {
                    DocComparePageNumberLabel.Content = (i + 1).ToString() + " / " + docCompareChildPanel1.Children.Count.ToString();
                    docCompareSideGridShown = i;
                    highlightSideGrid();
                    break;
                }
            }
        }

        private void updateDocSelectionComboBox()
        {
            // update combo box left
            ObservableCollection<string> items = new ObservableCollection<string>();
            for (int i = 0; i < documents.Count; i++)
            {
                if (documents[i].filePath != null && i != secondDocToShow)
                {
                    items.Add(Path.GetFileName(documents[i].filePath));
                }
            }
            Doc1NameLabelComboBox.ItemsSource = items;

            // update combo box right
            items = new ObservableCollection<string>();
            for (int i = 0; i < documents.Count; i++)
            {
                if (documents[i].filePath != null && i != firstDocToShow)
                {
                    items.Add(Path.GetFileName(documents[i].filePath));
                }
            }
            Doc2NameLabelComboBox.ItemsSource = items;
        }

        private void DocCompareDragDropZone2_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {

                var data = e.Data.GetData(DataFormats.FileDrop) as string[];

                if (data.Length == 1) // only one file drop
                {
                    if (documents[1].filePath == null)
                    {
                        documents[1].filePath = data[0];
                        documents[1].loaded = false;
                        documents[1].processed = false;
                    }
                    else
                    {
                        documents[2].filePath = data[0];
                        documents[2].loaded = false;
                        documents[2].processed = false;
                    }

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc2Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc2.Visibility = Visibility.Visible;

                        //updateDocSelectionComboBox();

                    });
                }
                else
                {
                    for (int i = 0; i < Math.Min(MAX_DOC_COUNT, data.Length); i++)
                    {
                        documents[i].filePath = data[i];
                        documents[i].loaded = false;
                        documents[i].processed = false;
                    }

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc1Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc1.Visibility = Visibility.Visible;
                        Doc2Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc2.Visibility = Visibility.Visible;

                        //updateDocSelectionComboBox();

                    });
                }

                threadDoc1 = new Thread(new ThreadStart(processDocThread));
                threadDoc1.Start();
            }
        }

        private void Doc2NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc2NameLabelComboBox.SelectedItem.ToString();
                secondDocToShow = documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayImageRight(secondDocToShow);
                updateDocSelectionComboBox();
            }
            catch
            {
                secondDocToShow = 1;
                DisplayImageRight(secondDocToShow);
                Doc2NameLabelComboBox.SelectedIndex = 0;
                updateDocSelectionComboBox();
            }
        }

        private void Doc1NameLabelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string fileName = Doc1NameLabelComboBox.SelectedItem.ToString();
                firstDocToShow = documents.FindIndex(x => Path.GetFileName(x.filePath) == fileName);
                DisplayImageLeft(firstDocToShow);
                updateDocSelectionComboBox();
            }
            catch
            {
                firstDocToShow = 0;
                DisplayImageLeft(firstDocToShow);
                Doc1NameLabelComboBox.SelectedIndex = 0;
                updateDocSelectionComboBox();
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

        private void CompareDocsThread()
        {
            try
            {
                docCompareRunning = true;
                Document.compareDocs(documents[firstDocToShow].imageFolder, documents[secondDocToShow].imageFolder, Path.Join(workingDir, "compare"), out pageIndices, out totalLen);
                documents[firstDocToShow].docCompareIndices = new List<int>();
                documents[secondDocToShow].docCompareIndices = new List<int>();

                if (totalLen != 0) // ? comparion successful
                {
                    for (int i = totalLen - 1; i >= 0; i--)
                    {
                        documents[firstDocToShow].docCompareIndices.Add((int)pageIndices[i]);
                        documents[secondDocToShow].docCompareIndices.Add((int)pageIndices[i + totalLen]);
                    }
                }

                docCompareRunning = false;

                Dispatcher.Invoke(() =>
                {
                    DisplayComparisonResult();
                    highlightSideGrid();
                    ProgressBarDocCompare.Visibility = Visibility.Hidden;
                });

            }
            catch
            {
                docCompareRunning = false;
            }
        }

        private void DisplayImageLeft(int docIndex)
        {
            Dispatcher.Invoke(() =>
            {
                if (documents[docIndex].filePath != null)
                {
                    int pageCounter = 0;
                    childPanel1 = new StackPanel();
                    childPanel1.Background = Brushes.Gray;
                    childPanel1.HorizontalAlignment = HorizontalAlignment.Stretch;

                    DirectoryInfo di = new DirectoryInfo(documents[docIndex].imageFolder);
                    FileInfo[] fi = di.GetFiles();

                    for (int i = 0; i < fi.Length; i++)
                    {
                        Image thisImage = new Image();

                        var stream = File.OpenRead(Path.Join(documents[docIndex].imageFolder, i.ToString() + ".jpg"));
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
                    //DocCompareColorZone1.Visibility = Visibility.Hidden;
                }
            });
        }

        private void DisplayImageRight(int docIndex)
        {
            Dispatcher.Invoke(() =>
            {
                if (documents[docIndex].filePath != null)
                {
                    int pageCounter = 0;
                    childPanel2 = new StackPanel();
                    childPanel2.Background = Brushes.Gray;
                    childPanel2.HorizontalAlignment = HorizontalAlignment.Stretch;

                    DirectoryInfo di = new DirectoryInfo(documents[docIndex].imageFolder);
                    FileInfo[] fi = di.GetFiles();

                    for (int i = 0; i < fi.Length; i++)
                    {
                        Image thisImage = new Image();
                        var stream = File.OpenRead(Path.Join(documents[docIndex].imageFolder, i.ToString() + ".jpg"));
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
                    //ProgressBarDoc2.Visibility = Visibility.Hidden;
                    //DocCompareColorZone2.Visibility = Visibility.Hidden
                }
            });
        }

        private void DisplayComparisonResult()
        {
            int pageCounter = 0;
            Dispatcher.Invoke(() =>
            {
                DocCompareNameLabel1.Content = Path.GetFileName(documents[firstDocToShow].filePath);
                DocCompareNameLabel2.Content = Path.GetFileName(documents[secondDocToShow].filePath);

                docCompareChildPanel1 = new StackPanel();
                docCompareChildPanel2 = new StackPanel();
                docCompareChildPanel1.Background = Brushes.Gray;
                docCompareChildPanel1.HorizontalAlignment = HorizontalAlignment.Stretch;
                docCompareChildPanel2.Background = Brushes.Gray;
                docCompareChildPanel2.HorizontalAlignment = HorizontalAlignment.Stretch;
                Image thisImage;
                FileStream stream;
                BitmapImage bitmap;

                for (int i = 0; i < totalLen; i++) // going through all the pages of the longest document
                {
                    Grid thisGrid = new Grid();
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc1
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc2

                    if (documents[firstDocToShow].docCompareIndices[i] != -1) // doc 1 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[firstDocToShow].imageFolder, documents[firstDocToShow].docCompareIndices[i].ToString() + ".jpg"));
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

                    if (documents[secondDocToShow].docCompareIndices[i] != -1) // doc 2 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[secondDocToShow].imageFolder, documents[secondDocToShow].docCompareIndices[i].ToString() + ".jpg"));
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

                        if (documents[firstDocToShow].docCompareIndices[i] != -1 && showMask == true)
                        {
                            if (File.Exists(Path.Join(workingDir, Path.Join("compare", documents[firstDocToShow].docCompareIndices[i].ToString() + "_" + documents[secondDocToShow].docCompareIndices[i].ToString() + ".png"))))
                            {
                                thisImage = new Image();
                                stream = File.OpenRead(Path.Join(workingDir, Path.Join("compare", documents[firstDocToShow].docCompareIndices[i].ToString() + "_" + documents[secondDocToShow].docCompareIndices[i].ToString() + ".png")));
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

                                //thisImage.Effect = new DropShadowEffect() { BlurRadius = 5, Color = Colors.Black, ShadowDepth = 0 };
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
                for (int i = 0; i < totalLen; i++) // going through all the pages of the longest document
                {
                    /*
                    Grid topGrid = new Grid();
                    topGrid.RowDefinitions.Add(new RowDefinition()); // doc1
                    topGrid.RowDefinitions.Add(new RowDefinition()); // page number
                    */
                    Border thisBorder = new Border();
                    thisBorder.BorderBrush = Brushes.Transparent;
                    thisBorder.BorderThickness = new Thickness(0);
                    Grid thisGrid = new Grid();
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Auto) }); // page number
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc1
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc2
                    thisGrid.MouseLeftButtonDown += (sen, ev) => { handleMouseClickOnSideScrollView(sen, ev); };
                    thisGrid.Name = "SideGrid" + i.ToString();
                    thisBorder.Child = thisGrid;
                    //Grid.SetRow(thisGrid, 0);

                    if (documents[firstDocToShow].docCompareIndices[i] != -1) // doc 1 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[firstDocToShow].imageFolder, documents[firstDocToShow].docCompareIndices[i].ToString() + ".jpg"));
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

                    if (documents[secondDocToShow].docCompareIndices[i] != -1) // doc 2 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[secondDocToShow].imageFolder, documents[secondDocToShow].docCompareIndices[i].ToString() + ".jpg"));
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
                        if (documents[firstDocToShow].docCompareIndices[i] != -1 && showMask == true)
                        {
                            if (File.Exists(Path.Join(workingDir, Path.Join("compare", documents[firstDocToShow].docCompareIndices[i].ToString() + "_" + documents[secondDocToShow].docCompareIndices[i].ToString() + ".png"))))
                            {
                                thisImage = new Image();
                                stream = File.OpenRead(Path.Join(workingDir, Path.Join("compare", documents[firstDocToShow].docCompareIndices[i].ToString() + "_" + documents[secondDocToShow].docCompareIndices[i].ToString() + ".png")));
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

                                //thisBorder.BorderBrush = Brushes.Red;
                                thisBorder.Background = new SolidColorBrush(Color.FromArgb(128, 255, 30, 30));
                                //thisBorder.BorderThickness = new Thickness(2);
                            }
                        }
                    }

                    pageCounter++;

                    Label thisLabel = new Label();
                    thisLabel.HorizontalAlignment = HorizontalAlignment.Center;
                    thisLabel.VerticalAlignment = VerticalAlignment.Center;
                    thisLabel.Content = (i + 1).ToString();
                    thisLabel.MinWidth = 20;
                    Grid.SetColumn(thisLabel, 0);

                    thisGrid.Children.Add(thisLabel);
                    //topGrid.Children.Add(thisGrid);
                    docCompareChildPanel2.Children.Add(thisBorder);
                }
                DocCompareSideScrollViewer.Content = docCompareChildPanel2;

                docCompareGrid.Visibility = Visibility.Visible;
            });
        }

        private void handleMouseClickOnSideScrollView(object sender, MouseButtonEventArgs e)
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
            highlightSideGrid();
        }

        private void highlightSideGrid()
        {
            double accuHeight = 0;
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
                        thisGrid.Background = Brushes.LightGray;

                        Size thisSize = thisGrid.DesiredSize;

                        DocCompareSideScrollViewer.ScrollToVerticalOffset(accuHeight);
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
