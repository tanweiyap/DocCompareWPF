using DocCompareWPF.Classes;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
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
        };

        List<Document> documents;
        string workingDir = Path.Join(Directory.GetCurrentDirectory(), "temp");
        bool doc1Loaded, doc2Loaded, doc1Processed, doc2Processed, docCompareRunning;
        ArrayList pageIndices;
        int totalLen;

        // Stack panel for viewing documents in scrollviewer control
        StackPanel childPanel1, childPanel2;
        Thread threadDoc1, threadDoc2, threadCompare;

        public MainWindow()
        {
            InitializeComponent();
            documents = new List<Document>();
            doc1Processed = true;
            doc2Processed = true;

            // Add dummy documents
            for (int i = 0; i < 2; i++)
            {
                documents.Add(new Document());
            }

            // create the temporary folders for converted images
            Directory.CreateDirectory(Path.Join(workingDir, "doc1"));
            Directory.CreateDirectory(Path.Join(workingDir, "doc2"));
            Directory.CreateDirectory(Path.Join(workingDir, "compare"));
            documents[0].imageFolder = Path.Join(workingDir, "doc1");
            documents[1].imageFolder = Path.Join(workingDir, "doc2");

            // GUI stuff
            DragDropPanel.Visibility = Visibility.Visible;
            DocComparePanel.Visibility = Visibility.Hidden;
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
            SetVisiblePanel(SidePanels.DOCCOMPARE);
            Application.Current.Dispatcher.Invoke(() =>
            {
                ProgressBarDocCompare.Visibility = Visibility.Visible;
            });

            if (doc1Loaded == true && doc2Loaded == true && docCompareRunning == false)
            {
                threadCompare = new Thread(new ThreadStart(CompareDocsThread));
                threadCompare.Start();
            }
        }

        private void SetVisiblePanel(SidePanels p_sidePanel)
        {
            //Brush brush = FindResource("SidePanelActiveBackground") as Brush;

            switch (p_sidePanel)
            {
                case SidePanels.DRAGDROP:
                    //SidePanelDocCompareBackground.Background = brush;
                    DragDropPanel.Visibility = Visibility.Visible;
                    DocComparePanel.Visibility = Visibility.Hidden;
                    break;
                case SidePanels.DOCCOMPARE:
                    //SidePanelDocCompareBackground.Background = brush;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    DocComparePanel.Visibility = Visibility.Visible;
                    break;
                default:
                    //SidePanelDocCompareBackground.Background = Brushes.Transparent;
                    DragDropPanel.Visibility = Visibility.Hidden;
                    break;
            }
        }

        private void DocCompareDragDropZone1_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                if(doc1Processed == true)
                    e.Effects = DragDropEffects.Copy;
                else
                    e.Effects = DragDropEffects.None;
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
                if (doc1Processed == true && docCompareRunning == false)
                {
                    var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                    // handle the files here!
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc1Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc1.Visibility = Visibility.Visible;
                    });

                    Doc1NameLabel.Content = Path.GetFileName(data[0]);
                    doc1Loaded = false;
                    doc1Processed = false;
                    threadDoc1 = new Thread(new ParameterizedThreadStart(processDoc1Thread));
                    threadDoc1.Start(data[0]);
                }
            }
        }

        private void processDoc1Thread(object data)
        {
            Dispatcher.Invoke(() =>
            {
                DocCompareScrollViewer1.Content = null;
            });

            documents[0].clearFolder();
            documents[0].filePath = (string)data;
            documents[0].detectFileType();
            documents[0].docID = "doc1";
            //DocCompareDropZone1Label.Content = documents[0].fileType.ToString();

            if (documents[0].fileType == Document.FileTypes.PDF)
            {
                if (documents[0].readPDF() == 0)
                {
                    doc1Loaded = true;
                }
            }

            if (doc1Loaded == true)
            {
                DisplayImage(0);
            }
            
        }

        private void CloseDoc1Button_Click(object sender, RoutedEventArgs e)
        {
            documents[0] = new Document();
            documents[0].imageFolder = Path.Join(workingDir, "doc1");
            doc1Loaded = false;
            doc1Processed = true;
            Doc1Grid.Visibility = Visibility.Hidden;
        }

        private void CloseDoc2Button_Click(object sender, RoutedEventArgs e)
        {
            documents[1] = new Document();
            documents[1].imageFolder = Path.Join(workingDir, "doc2");
            doc2Loaded = false;
            doc2Processed = true;
            Doc2Grid.Visibility = Visibility.Hidden;
        }

        private void DocCompareDragDropZone2_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                if (doc2Processed == true && docCompareRunning == false)
                    e.Effects = DragDropEffects.Copy;
                else
                    e.Effects = DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void CloseDocCompareButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void DocCompareDragDropZone2_Drop(object sender, DragEventArgs e)
        {
            if (null != e.Data && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                if (doc2Processed == true)
                {
                    var data = e.Data.GetData(DataFormats.FileDrop) as string[];
                    // handle the files here!
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Doc2Grid.Visibility = Visibility.Hidden;
                        ProgressBarDoc2.Visibility = Visibility.Visible;
                    });

                    doc2Loaded = false;
                    doc2Processed = false;
                    Doc2NameLabel.Content = Path.GetFileName(data[0]);
                    threadDoc2 = new Thread(new ParameterizedThreadStart(processDoc2Thread));
                    threadDoc2.Start(data[0]);
                }
            }
        }

        private void processDoc2Thread(object data)
        {
            Dispatcher.Invoke(() =>
            {
                DocCompareScrollViewer2.Content = null;
            });

            documents[1].clearFolder();
            documents[1].filePath = (string)data;
            documents[1].detectFileType();
            documents[1].docID = "doc1";
            //DocCompareDropZone1Label.Content = documents[0].fileType.ToString();

            if (documents[1].fileType == Document.FileTypes.PDF)
            {
                if (documents[1].readPDF() == 0)
                {
                    doc2Loaded = true;
                }
            }

            if (doc2Loaded == true)
            {
                DisplayImage(1);
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
                Document.compareDocs(documents[0].imageFolder, documents[1].imageFolder, Path.Join(workingDir, "compare"), out pageIndices, out totalLen);
                documents[0].docCompareIndices = new List<int>();
                documents[1].docCompareIndices = new List<int>();

                if (totalLen != 0) // ? comparion successful
                {
                    for(int i = totalLen - 1; i >= 0; i--)
                    {
                        documents[0].docCompareIndices.Add((int)pageIndices[i]);
                        documents[1].docCompareIndices.Add((int)pageIndices[i + totalLen]);
                    }
                }

                docCompareRunning = false;

                Dispatcher.Invoke(() =>
                {
                    DisplayComparisonResult();
                    ProgressBarDocCompare.Visibility = Visibility.Hidden;
                });

            }catch
            {
                docCompareRunning = false;
            }
        }

        private void DisplayImage(int docIndex)
        {
            Dispatcher.Invoke(() =>
            {
                int pageCounter = 0;
                if (docIndex == 0)
                {
                    childPanel1 = new StackPanel();
                    childPanel1.Background = Brushes.Gray;
                    childPanel1.HorizontalAlignment = HorizontalAlignment.Stretch;

                    DirectoryInfo di = new DirectoryInfo(documents[0].imageFolder);

                    foreach (FileInfo file in di.EnumerateFiles())
                    {
                        Image thisImage = new Image();

                        var stream = File.OpenRead(file.FullName);
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
                    doc1Processed = true;
                    ProgressBarDoc1.Visibility = Visibility.Hidden;
                    //DocCompareColorZone1.Visibility = Visibility.Hidden;
                }

                if (docIndex == 1)
                {
                    childPanel2 = new StackPanel();
                    childPanel2.Background = Brushes.Gray;
                    childPanel2.HorizontalAlignment = HorizontalAlignment.Stretch;

                    DirectoryInfo di = new DirectoryInfo(documents[1].imageFolder);

                    foreach (FileInfo file in di.EnumerateFiles())
                    {
                        Image thisImage = new Image();
                        var stream = File.OpenRead(file.FullName);
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
                    doc2Processed = true;
                    //DocCompareColorZone2.Visibility = Visibility.Hidden;
                }
            });
        }

        private void DisplayComparisonResult()
        {
            int pageCounter = 0;
            Dispatcher.Invoke(() =>
            {
                StackPanel childPanel = new StackPanel();
                StackPanel childPanel2 = new StackPanel();
                childPanel.Background = Brushes.Gray;
                childPanel.HorizontalAlignment = HorizontalAlignment.Stretch;
                childPanel2.Background = Brushes.Gray;
                childPanel2.HorizontalAlignment = HorizontalAlignment.Stretch;
                Image thisImage;
                FileStream stream;
                BitmapImage bitmap;

                for (int i = 0; i < totalLen; i++) // going through all the pages of the longest document
                {
                    Grid thisGrid = new Grid();
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc1
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc2

                    if(documents[0].docCompareIndices[i] != -1) // doc 1 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[0].imageFolder, documents[0].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
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
                        Grid.SetColumn(thisImage, 0);
                        thisGrid.Children.Add(thisImage);
                    }

                    if (documents[1].docCompareIndices[i] != -1) // doc 2 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[1].imageFolder, documents[1].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
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
                        Grid.SetColumn(thisImage, 1);
                        thisGrid.Children.Add(thisImage);
                    }

                    pageCounter++;
                    childPanel.Children.Add(thisGrid);
                }

                DocCompareMainScrollViewer.Content = childPanel;

                for (int i = 0; i < totalLen; i++) // going through all the pages of the longest document
                {
                    Grid thisGrid = new Grid();
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc1
                    thisGrid.ColumnDefinitions.Add(new ColumnDefinition()); // doc2

                    if (documents[0].docCompareIndices[i] != -1) // doc 1 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[0].imageFolder, documents[0].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
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
                        Grid.SetColumn(thisImage, 0);
                        thisGrid.Children.Add(thisImage);
                    }

                    if (documents[1].docCompareIndices[i] != -1) // doc 2 has a valid page
                    {
                        thisImage = new Image();
                        stream = File.OpenRead(Path.Join(documents[1].imageFolder, documents[1].docCompareIndices[i].ToString() + ".jpg"));
                        bitmap = new BitmapImage();
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
                        Grid.SetColumn(thisImage, 1);
                        thisGrid.Children.Add(thisImage);
                    }

                    pageCounter++;
                    childPanel2.Children.Add(thisGrid);
                }
                DocCompareSideScrollViewer.Content = childPanel2;
            });
        }
    }
}
