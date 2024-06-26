﻿using DocCompareWPF.Classes;
using System;
using System.Windows.Media;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Input;
using System.Windows.Controls;
using System.Collections.Generic;

namespace DocCompareWPF
{
    /// <summary>
    /// Interaktionslogik für NoCompareZoneWindow.xaml
    /// </summary>
    public partial class NoCompareZoneWindow : Window
    {
        int pageIndex;
        Document localDoc;
        bool isMouseDown;
        EditMode mode;
        Point startPoint, endPoint;
        List<List<double>> rects;
        public List<List<int>> rectsFinal;
        List<List<int>> rectsInput;
        int selectedRectIndex = -1;
        List<double> newPos;
        EdgeType lastEdge = EdgeType.NONE;
        string pathToImg = "";
        string pathToFolder = "";

        private enum EditMode
        {
            SELECT,
            CREATE,
            DELETE
        };

        public NoCompareZoneWindow()
        {
            InitializeComponent();

            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            SelectButtonBackground.Background = brush;
            CreateRectBackground.Background = Brushes.Transparent;
            mode = EditMode.SELECT;
            rects = new List<List<double>>();
            rectsFinal = new List<List<int>>();

        }

        public void SetupWindow(Document doc, List<List<int>> _rects, string _path)
        {
            if (doc.filePath != null)
            {
                localDoc = doc;
                DirectoryInfo di = new DirectoryInfo(localDoc.imageFolder);
                FileInfo[] fi = di.GetFiles();

                NumberOfPagesLabel.Content = "/ " + fi.Length;
                SelectedPageTextBox.Text = "1";
                pageIndex = 0;

                BitmapImage bi = new BitmapImage();
                bi.BeginInit();
                bi.UriSource = new Uri(fi[0].FullName);
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.EndInit();

                PageImage.Source = bi;

                if (pageIndex == fi.Length - 1)
                {
                    NextPageButton.IsEnabled = false;
                }

                PrevPageButton.IsEnabled = false;
            }

            if (_rects.Count != 0)
            {
                rectsInput = _rects;
            }

            pathToFolder = _path;
            pathToImg = Path.Join(_path, "noCompareMask" + Guid.NewGuid().ToString() + ".png");
            Directory.CreateDirectory(_path);
        }

        private void CreateRectButton_Click(object sender, RoutedEventArgs e)
        {
            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            SelectButtonBackground.Background = Brushes.Transparent;
            CreateRectBackground.Background = brush;
            DeleteRectBackground.Background = Brushes.Transparent;
            mode = EditMode.CREATE;
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            SelectButtonBackground.Background = brush;
            CreateRectBackground.Background = Brushes.Transparent;
            DeleteRectBackground.Background = Brushes.Transparent;
            mode = EditMode.SELECT;
        }

        private void PrevPageButton_Click(object sender, RoutedEventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(localDoc.imageFolder);
            FileInfo[] fi = di.GetFiles();

            pageIndex--;
            SelectedPageTextBox.Text = (pageIndex + 1).ToString();

            BitmapImage bi = new BitmapImage();
            bi.BeginInit();
            bi.UriSource = new Uri(fi[pageIndex].FullName);
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.EndInit();

            PageImage.Source = bi;

            if (pageIndex == fi.Length - 1)
            {
                NextPageButton.IsEnabled = false;
            }
            else
            {
                NextPageButton.IsEnabled = true;
            }

            if (pageIndex == 0)
            {
                PrevPageButton.IsEnabled = false;
            }
        }

        private void NextPageButton_Click(object sender, RoutedEventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(localDoc.imageFolder);
            FileInfo[] fi = di.GetFiles();

            pageIndex++;
            SelectedPageTextBox.Text = (pageIndex + 1).ToString();

            BitmapImage bi = new BitmapImage();
            bi.BeginInit();
            bi.UriSource = new Uri(fi[pageIndex].FullName);
            bi.CacheOption = BitmapCacheOption.OnLoad;
            bi.EndInit();

            PageImage.Source = bi;

            if (pageIndex == fi.Length - 1)
            {
                NextPageButton.IsEnabled = false;
            }

            if (pageIndex >= 0)
            {
                PrevPageButton.IsEnabled = true;
            }
        }

        private void DoneButton_Click(object sender, RoutedEventArgs e)
        {
            double actualW = (PageImage.Source as BitmapImage).PixelWidth;
            double actualH = (PageImage.Source as BitmapImage).PixelHeight;

            double currW = PageCanvas.ActualWidth;
            double currH = PageCanvas.ActualHeight;

            foreach (List<double> l in rects)
            {
                l[0] = l[0] * actualW / currW;
                l[2] = l[2] * actualW / currW;
                l[1] = l[1] * actualH / currH;
                l[3] = l[3] * actualH / currH;

                List<int> temp = new List<int>();
                temp.Add((int)l[0]);
                temp.Add((int)l[1]);
                temp.Add((int)l[2]);
                temp.Add((int)l[3]);
                rectsFinal.Add(temp);
            }

            VisualBrush brush = FindResource("HatchBrush") as VisualBrush;

            for (int i = 0; i < PageCanvas.Children.Count; i++)
            {
                (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).Fill = brush;
                //(PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).Opacity = 1.0;
                (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).StrokeThickness = 1;
            }

            // Save current canvas transform
            Transform transform = PageCanvas.LayoutTransform;
            // reset current transform (in case it is scaled or rotated)
            PageCanvas.LayoutTransform = null;

            // Get the size of canvas
            Size size = new Size(PageCanvas.ActualWidth, PageCanvas.ActualHeight);
            // Measure and arrange the PageCanvas
            // VERY IMPORTANT
            PageCanvas.Measure(size);
            PageCanvas.Arrange(new Rect(size));

            // Create a render bitmap and push the PageCanvas to it
            RenderTargetBitmap renderBitmap =
              new RenderTargetBitmap(
                (int)size.Width,
                (int)size.Height,
                96d,
                96d,
                PixelFormats.Pbgra32);
            renderBitmap.Render(PageCanvas);

            // delete previous            
            DirectoryInfo di = new DirectoryInfo(pathToFolder);
            FileInfo[] fi = di.GetFiles();
            if (fi.Length != 0)
            {
                foreach (FileInfo f in fi)
                    f.Delete();
            }

            // Create a file stream for saving image only if there are no compare zones
            if (rects.Count != 0)
            {
                using (FileStream outStream = new FileStream(pathToImg, FileMode.Create))
                {
                    // Use png encoder for our data
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    // push the rendered bitmap to it
                    encoder.Frames.Add(BitmapFrame.Create(renderBitmap));
                    // save the data to the stream
                    encoder.Save(outStream);
                }
            }

            // Restore previously saved layout
            PageCanvas.LayoutTransform = transform;

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void PageCanvas_MouseEnter(object sender, MouseEventArgs e)
        {
            if (mode == EditMode.CREATE)
            {
                PageCanvas.Cursor = Cursors.Cross;
            }
            else
            {
                PageCanvas.Cursor = Cursors.Arrow;
            }
        }

        private void PageCanvas_MouseLeave(object sender, MouseEventArgs e)
        {
            PageCanvas.Cursor = Cursors.Arrow;
        }

        private void PageCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (mode == EditMode.CREATE)
            {
                isMouseDown = true;
                startPoint = e.GetPosition(PageCanvas);

                System.Windows.Shapes.Rectangle newRect = new System.Windows.Shapes.Rectangle();
                newRect.Stroke = Brushes.Black;

                Brush brush = FindResource("WarningBackgroundBrush") as Brush;

                newRect.Fill = brush;
                newRect.Opacity = 1.0;
                newRect.Width = 1;
                newRect.Height = 1;

                Canvas.SetLeft(newRect, startPoint.X);
                Canvas.SetTop(newRect, startPoint.Y);
                PageCanvas.Children.Add(newRect);
            }
            else if (mode == EditMode.SELECT && selectedRectIndex != -1 && PageCanvas.Cursor != Cursors.Arrow)
            {
                startPoint = e.GetPosition(PageCanvas);
                newPos = new List<double>();
                newPos.Add(rects[selectedRectIndex][0]);
                newPos.Add(rects[selectedRectIndex][1]);
                newPos.Add(rects[selectedRectIndex][2]);
                newPos.Add(rects[selectedRectIndex][3]);

                isMouseDown = true;
            }
            else if (mode == EditMode.SELECT)
            {
                selectedRectIndex = FindRect(e.GetPosition(PageCanvas));
                if (selectedRectIndex != -1)
                {
                    /*
                    System.Windows.Shapes.Rectangle newRect = new System.Windows.Shapes.Rectangle();
                    newRect.Stroke = Brushes.Black;
                    */
                    /*
                    newRect.Fill = brush;
                    newRect.Opacity = 0.5;
                    newRect.Width = Math.Abs(rects[selectedRectIndex][2]- rects[selectedRectIndex][0]);
                    newRect.Height = Math.Abs(rects[selectedRectIndex][3] - rects[selectedRectIndex][1]);

                    Canvas.SetLeft(newRect, rects[selectedRectIndex][0]);
                    Canvas.SetTop(newRect, rects[selectedRectIndex][1]);
                    PageCanvas.Children.RemoveAt(selectedRectIndex);
                    PageCanvas.Children.Insert(selectedRectIndex, newRect);
                    */

                    // reset all colors
                    VisualBrush brush = FindResource("HatchBrush") as VisualBrush;

                    for (int i = 0; i < PageCanvas.Children.Count; i++)
                    {
                        (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).Fill = brush;
                        (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).StrokeThickness = 1;
                    }

                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Fill = brush;
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).StrokeThickness = 3;

                }
                else // no rectangle selected
                {
                    VisualBrush brush = FindResource("HatchBrush") as VisualBrush;
                    for (int i = 0; i < PageCanvas.Children.Count; i++)
                    {
                        (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).Fill = brush;
                        (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).StrokeThickness = 1;
                    }
                }

            }
            else if (mode == EditMode.DELETE)
            {
                selectedRectIndex = FindRect(e.GetPosition(PageCanvas));
                if (selectedRectIndex != -1)
                {
                    PageCanvas.Children.RemoveAt(selectedRectIndex);
                    rects.RemoveAt(selectedRectIndex);
                    selectedRectIndex = -1;
                    isMouseDown = false;

                    Brush brush = FindResource("SecondaryAccentBrush") as Brush;
                    SelectButtonBackground.Background = brush;
                    CreateRectBackground.Background = Brushes.Transparent;
                    DeleteRectBackground.Background = Brushes.Transparent;
                    mode = EditMode.SELECT;
                    PageCanvas.Cursor = Cursors.Arrow;
                }
            }

        }

        private void PageCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (mode == EditMode.CREATE && isMouseDown == true)
            {
                endPoint = e.GetPosition(PageCanvas);
                System.Windows.Shapes.Rectangle newRect = new System.Windows.Shapes.Rectangle();
                newRect.Stroke = Brushes.Black;

                VisualBrush brush = FindResource("HatchBrush") as VisualBrush;

                newRect.Fill = brush;
                newRect.Opacity = 1.0;
                newRect.Width = Math.Abs(endPoint.X - startPoint.X);
                newRect.Height = Math.Abs(endPoint.Y - startPoint.Y);

                Canvas.SetLeft(newRect, Math.Min(startPoint.X, endPoint.X));
                Canvas.SetTop(newRect, Math.Min(startPoint.Y, endPoint.Y));
                PageCanvas.Children.RemoveAt(PageCanvas.Children.Count - 1);
                PageCanvas.Children.Add(newRect);

                if (System.Windows.Input.Mouse.LeftButton == MouseButtonState.Released)
                {
                    isMouseDown = false;
                    List<double> values = new List<double>();
                    values.Add(Math.Min(startPoint.X, endPoint.X));
                    values.Add(Math.Min(startPoint.Y, endPoint.Y));
                    values.Add(Math.Max(startPoint.X, endPoint.X));
                    values.Add(Math.Max(startPoint.Y, endPoint.Y));
                    rects.Add(values);

                    Brush brush2 = FindResource("SecondaryAccentBrush") as Brush;
                    SelectButtonBackground.Background = brush2;
                    CreateRectBackground.Background = Brushes.Transparent;
                    mode = EditMode.SELECT;
                    PageCanvas.Cursor = Cursors.Arrow;
                }
            }
            else if (mode == EditMode.SELECT && isMouseDown == true)
            {
                if (Mouse.LeftButton == MouseButtonState.Released)
                {
                    isMouseDown = false;
                    rects.RemoveAt(selectedRectIndex);
                    rects.Insert(selectedRectIndex, newPos);
                    lastEdge = EdgeType.NONE;
                    //selectedRectIndex = -1;
                    return;
                }

                if (PageCanvas.Cursor == Cursors.SizeAll)
                {
                    endPoint = e.GetPosition(PageCanvas);

                    newPos[0] = newPos[0] - (startPoint.X - endPoint.X);
                    newPos[2] = newPos[2] - (startPoint.X - endPoint.X);
                    newPos[1] = newPos[1] - (startPoint.Y - endPoint.Y);
                    newPos[3] = newPos[3] - (startPoint.Y - endPoint.Y);

                    startPoint = endPoint;

                    if (newPos[0] < 0)
                    {
                        newPos[0] = 0;
                        newPos[2] = (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Width;
                    }

                    if (newPos[2] > PageCanvas.ActualWidth)
                    {
                        newPos[0] = PageCanvas.ActualWidth - (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Width;
                        newPos[2] = PageCanvas.ActualWidth;
                    }

                    if (newPos[1] < 0)
                    {
                        newPos[1] = 0;
                        newPos[3] = (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Height;
                    }

                    if (newPos[3] > PageCanvas.ActualHeight)
                    {
                        newPos[1] = PageCanvas.ActualHeight - (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Height;
                        newPos[3] = PageCanvas.ActualHeight;
                    }

                    Canvas.SetLeft((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[0]);
                    Canvas.SetTop((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[1]);
                }
                else if (PageCanvas.Cursor == Cursors.SizeNS) // resize top or bottom
                {
                    endPoint = e.GetPosition(PageCanvas);

                    if (lastEdge == EdgeType.NONE)
                    {
                        lastEdge = FindRectEdge(endPoint);
                    }

                    if (lastEdge == EdgeType.TOP)
                    {
                        if (newPos[1] - (startPoint.Y - endPoint.Y) <= newPos[3] - 15)
                            newPos[1] = newPos[1] - (startPoint.Y - endPoint.Y);
                    }
                    else if (lastEdge == EdgeType.BOTTOM)
                    {
                        if (newPos[3] - (startPoint.Y - endPoint.Y) >= newPos[1] + 15)
                            newPos[3] = newPos[3] - (startPoint.Y - endPoint.Y);
                    }

                    if (lastEdge == EdgeType.BOTTOM)
                    {
                        if (newPos[3] > PageCanvas.ActualHeight)
                        {
                            newPos[3] = PageCanvas.ActualHeight;
                        }

                        if (newPos[3] <= newPos[1])
                        {
                            newPos[3] = newPos[1] + 15;
                        }
                    }
                    else
                    {
                        if (newPos[1] < 0)
                        {
                            newPos[1] = 0;
                        }

                        if (newPos[1] >= newPos[3])
                        {
                            newPos[1] = newPos[3] - 15;
                        }
                    }

                    startPoint = endPoint;

                    Canvas.SetLeft((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[0]);
                    Canvas.SetTop((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[1]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Width = Math.Abs(newPos[2] - newPos[0]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Height = Math.Abs(newPos[3] - newPos[1]);
                    rects.RemoveAt(selectedRectIndex);
                    rects.Insert(selectedRectIndex, newPos);
                }
                else if (PageCanvas.Cursor == Cursors.SizeWE) // resize left or right
                {
                    endPoint = e.GetPosition(PageCanvas);
                    if (lastEdge == EdgeType.NONE)
                    {
                        lastEdge = FindRectEdge(endPoint);
                    }

                    if (lastEdge == EdgeType.LEFT)
                    {
                        if (newPos[0] - (startPoint.X - endPoint.X) <= newPos[2] - 15)
                            newPos[0] = newPos[0] - (startPoint.X - endPoint.X);
                    }
                    else if (lastEdge == EdgeType.RIGHT)
                    {
                        if (newPos[2] - (startPoint.X - endPoint.X) >= newPos[0] + 15)
                            newPos[2] = newPos[2] - (startPoint.X - endPoint.X);
                    }

                    if (lastEdge == EdgeType.RIGHT)
                    {
                        if (newPos[2] > PageCanvas.ActualWidth)
                        {
                            newPos[2] = PageCanvas.ActualWidth;
                        }

                        if (newPos[2] <= newPos[0])
                        {
                            newPos[2] = newPos[0] + 15;
                            lastEdge = EdgeType.NONE;
                        }
                    }
                    else
                    {
                        if (newPos[0] < 0)
                        {
                            newPos[0] = 0;
                        }

                        if (newPos[0] >= newPos[2])
                        {
                            newPos[0] = newPos[2] - 15;
                        }
                    }
                    startPoint = endPoint;

                    Canvas.SetLeft((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[0]);
                    Canvas.SetTop((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[1]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Width = Math.Abs(newPos[2] - newPos[0]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Height = Math.Abs(newPos[3] - newPos[1]);
                    rects.RemoveAt(selectedRectIndex);
                    rects.Insert(selectedRectIndex, newPos);
                }
                else if (PageCanvas.Cursor == Cursors.SizeNESW) // resize topright or bottomleft
                {
                    endPoint = e.GetPosition(PageCanvas);
                    if (lastEdge == EdgeType.NONE)
                    {
                        lastEdge = FindRectEdge(endPoint);
                    }

                    if (lastEdge == EdgeType.TOPRIGHT)
                    {
                        if(newPos[2] - (startPoint.X - endPoint.X) >= newPos[0] + 15)
                            newPos[2] = newPos[2] - (startPoint.X - endPoint.X);
                        if(newPos[1] - (startPoint.Y - endPoint.Y) <= newPos[3] - 15)
                            newPos[1] = newPos[1] - (startPoint.Y - endPoint.Y);
                    }
                    else if (lastEdge == EdgeType.BOTTOMLEFT)
                    {
                        if (newPos[0] - (startPoint.X - endPoint.X) <= newPos[2] - 15)
                            newPos[0] = newPos[0] - (startPoint.X - endPoint.X);
                        if (newPos[3] - (startPoint.Y - endPoint.Y) >= newPos[1] + 15)
                            newPos[3] = newPos[3] - (startPoint.Y - endPoint.Y);
                    }

                    if (lastEdge == EdgeType.BOTTOMLEFT)
                    {
                        if (newPos[3] > PageCanvas.ActualHeight)
                        {
                            newPos[3] = PageCanvas.ActualHeight;
                        }

                        if (newPos[3] <= newPos[1])
                        {
                            newPos[3] = newPos[1] + 15;
                        }

                        if (newPos[2] > PageCanvas.ActualWidth)
                        {
                            newPos[2] = PageCanvas.ActualWidth;
                        }

                        if (newPos[2] <= newPos[0])
                        {
                            newPos[2] = newPos[0] + 15;
                        }
                    }
                    else
                    {
                        if (newPos[1] < 0)
                        {
                            newPos[1] = 0;
                        }

                        if (newPos[1] >= newPos[3])
                        {
                            newPos[1] = newPos[3] - 15;
                        }

                        if (newPos[0] < 0)
                        {
                            newPos[0] = 0;
                        }

                        if (newPos[0] >= newPos[2])
                        {
                            newPos[0] = newPos[2] - 15;
                        }
                    }
                    startPoint = endPoint;

                    Canvas.SetLeft((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[0]);
                    Canvas.SetTop((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[1]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Width = Math.Abs(newPos[2] - newPos[0]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Height = Math.Abs(newPos[3] - newPos[1]);
                    rects.RemoveAt(selectedRectIndex);
                    rects.Insert(selectedRectIndex, newPos);
                }
                if (PageCanvas.Cursor == Cursors.SizeNWSE) // resize topleft or bottomright
                {
                    endPoint = e.GetPosition(PageCanvas);
                    if (lastEdge == EdgeType.NONE)
                    {
                        lastEdge = FindRectEdge(endPoint);
                    }

                    if (lastEdge == EdgeType.TOPLEFT)
                    {
                        if (newPos[0] - (startPoint.X - endPoint.X) <= newPos[2] - 15)
                            newPos[0] = newPos[0] - (startPoint.X - endPoint.X);
                        if (newPos[1] - (startPoint.Y - endPoint.Y) <= newPos[3] - 15)
                            newPos[1] = newPos[1] - (startPoint.Y - endPoint.Y);
                    }
                    else if (lastEdge == EdgeType.BOTTOMRIGHT)
                    {
                        if (newPos[2] - (startPoint.X - endPoint.X) >= newPos[0] + 15)
                            newPos[2] = newPos[2] - (startPoint.X - endPoint.X);
                        if (newPos[3] - (startPoint.Y - endPoint.Y) >= newPos[1] + 15)
                            newPos[3] = newPos[3] - (startPoint.Y - endPoint.Y);
                    }

                    if(lastEdge == EdgeType.TOPLEFT)
                    {
                        if (newPos[0] < 0)
                        {
                            newPos[0] = 0;
                        }

                        if (newPos[0] >= newPos[2])
                        {
                            newPos[0] = newPos[2] - 15;
                        }

                        if (newPos[1] < 0)
                        {
                            newPos[1] = 0;
                        }

                        if (newPos[1] >= newPos[3])
                        {
                            newPos[1] = newPos[3] - 15;
                        }
                    }
                    else
                    {
                        if (newPos[3] > PageCanvas.ActualHeight)
                        {
                            newPos[3] = PageCanvas.ActualHeight;
                        }

                        if (newPos[3] <= newPos[1])
                        {
                            newPos[3] = newPos[1] + 15;
                        }

                        if (newPos[2] > PageCanvas.ActualWidth)
                        {
                            newPos[2] = PageCanvas.ActualWidth;
                        }

                        if (newPos[2] <= newPos[0])
                        {
                            newPos[2] = newPos[0] + 15;
                        }
                    }

                    startPoint = endPoint;

                    Canvas.SetLeft((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[0]);
                    Canvas.SetTop((PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle), newPos[1]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Width = Math.Abs(newPos[2] - newPos[0]);
                    (PageCanvas.Children[selectedRectIndex] as System.Windows.Shapes.Rectangle).Height = Math.Abs(newPos[3] - newPos[1]);
                    rects.RemoveAt(selectedRectIndex);
                    rects.Insert(selectedRectIndex, newPos);
                }
                /*
                PageCanvas.Children.RemoveAt(selectedRectIndex);
                PageCanvas.Children.Insert(selectedRectIndex, newRect);
                */

            }

            if (mode == EditMode.SELECT && isMouseDown == false)
            {
                int foundRect = FindRect(e.GetPosition(PageCanvas));
                if (foundRect >= 0 && foundRect == selectedRectIndex)
                {
                    switch (FindRectEdge(e.GetPosition(PageCanvas)))
                    {
                        case EdgeType.TOP:
                        case EdgeType.BOTTOM:
                            PageCanvas.Cursor = Cursors.SizeNS;
                            break;
                        case EdgeType.LEFT:
                        case EdgeType.RIGHT:
                            PageCanvas.Cursor = Cursors.SizeWE;
                            break;
                        case EdgeType.TOPLEFT:
                        case EdgeType.BOTTOMRIGHT:
                            PageCanvas.Cursor = Cursors.SizeNWSE;
                            break;
                        case EdgeType.TOPRIGHT:
                        case EdgeType.BOTTOMLEFT:
                            PageCanvas.Cursor = Cursors.SizeNESW;
                            break;
                        default:
                            PageCanvas.Cursor = Cursors.SizeAll;
                            break;
                    }
                }
                else
                {
                    PageCanvas.Cursor = Cursors.Arrow;
                }
            }

            if (mode == EditMode.DELETE)
            {
                int foundRect = FindRect(e.GetPosition(PageCanvas));
                if (foundRect != -1)
                {
                    PageCanvas.Cursor = Cursors.Cross;
                }
                else
                {
                    PageCanvas.Cursor = Cursors.Arrow;
                }
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape && mode == EditMode.CREATE)
            {
                Brush brush = FindResource("SecondaryAccentBrush") as Brush;
                SelectButtonBackground.Background = brush;
                CreateRectBackground.Background = Brushes.Transparent;
                mode = EditMode.SELECT;
                PageCanvas.Cursor = Cursors.Arrow;

                if (isMouseDown == true)
                {
                    isMouseDown = false;
                    PageCanvas.Children.RemoveAt(PageCanvas.Children.Count - 1);
                }
            }
            else if (e.Key == Key.Escape && mode == EditMode.SELECT)
            {
                VisualBrush brush = FindResource("HatchBrush") as VisualBrush;
                for (int i = 0; i < PageCanvas.Children.Count; i++)
                {
                    (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).Fill = brush;
                    (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).StrokeThickness = 1;
                }
            }
            else if (e.Key == Key.Delete && mode == EditMode.SELECT && selectedRectIndex != -1)
            {
                PageCanvas.Children.RemoveAt(selectedRectIndex);
                rects.RemoveAt(selectedRectIndex);
                selectedRectIndex = -1;
                isMouseDown = false;
            }
        }

        private int FindRect(Point pos)
        {
            int ret = -1;

            for (int i = 0; i < rects.Count; i++)
            {
                if (rects[i][0] <= pos.X && rects[i][2] >= pos.X &&
                    rects[i][1] <= pos.Y && rects[i][3] >= pos.Y)
                {
                    return i;
                }
            }

            return ret;
        }

        private enum EdgeType
        {
            TOP,
            BOTTOM,
            LEFT,
            RIGHT,
            TOPLEFT,
            TOPRIGHT,
            BOTTOMLEFT,
            BOTTOMRIGHT,
            NONE,
        }

        private void DeleteRectButton_Click(object sender, RoutedEventArgs e)
        {
            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            SelectButtonBackground.Background = Brushes.Transparent;
            CreateRectBackground.Background = Brushes.Transparent;
            DeleteRectBackground.Background = brush;
            mode = EditMode.DELETE;

            VisualBrush brush2 = FindResource("HatchBrush") as VisualBrush;

            for (int i = 0; i < PageCanvas.Children.Count; i++)
            {
                (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).Fill = brush2;
                (PageCanvas.Children[i] as System.Windows.Shapes.Rectangle).StrokeThickness = 1;
            }
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            if (rectsInput != null)
            {
                foreach (List<int> i in rectsInput)
                {
                    double actualW = (PageImage.Source as BitmapImage).PixelWidth;
                    double actualH = (PageImage.Source as BitmapImage).PixelHeight;

                    double currW = PageCanvas.ActualWidth;
                    double currH = PageCanvas.ActualHeight;

                    List<double> temp = new List<double>();
                    temp.Add((double)i[0] * (double)currH / (double)actualH);
                    temp.Add((double)i[1] * (double)currH / (double)actualH);
                    temp.Add((double)i[2] * (double)currH / (double)actualH);
                    temp.Add((double)i[3] * (double)currH / (double)actualH);
                    rects.Add(temp);

                    VisualBrush brush = FindResource("HatchBrush") as VisualBrush;
                    System.Windows.Shapes.Rectangle newRect = new System.Windows.Shapes.Rectangle();
                    newRect.Stroke = Brushes.Black;
                    newRect.Fill = brush;
                    newRect.Opacity = 1.0;
                    newRect.Width = Math.Abs(temp[2] - temp[0]);
                    newRect.Height = Math.Abs(temp[3] - temp[1]);

                    Canvas.SetLeft(newRect, temp[0]);
                    Canvas.SetTop(newRect, temp[1]);
                    PageCanvas.Children.Add(newRect);
                }
            }

        }

        private EdgeType FindRectEdge(Point pos)
        {
            List<double> rectVertices = rects[selectedRectIndex];

            // check if TOP
            if (pos.X >= rectVertices[0] && pos.X <= rectVertices[2])
            {
                if (pos.Y >= rectVertices[1] && pos.Y <= rectVertices[1] + 5)
                {
                    if (pos.X >= rectVertices[0] && pos.X <= rectVertices[0] + 5)
                        return EdgeType.TOPLEFT;
                    else if (pos.X <= rectVertices[2] && pos.X >= rectVertices[2] - 5)
                        return EdgeType.TOPRIGHT;
                    else
                        return EdgeType.TOP;
                }
            }

            // check if Bottom
            if (pos.X >= rectVertices[0] && pos.X <= rectVertices[2])
            {
                if (pos.Y <= rectVertices[3] && pos.Y >= rectVertices[3] - 5)
                {
                    if (pos.X >= rectVertices[0] && pos.X <= rectVertices[0] + 5)
                        return EdgeType.BOTTOMLEFT;
                    else if (pos.X <= rectVertices[2] && pos.X >= rectVertices[2] - 5)
                        return EdgeType.BOTTOMRIGHT;
                    else
                        return EdgeType.BOTTOM;
                }
            }

            // check if Left
            if (pos.X >= rectVertices[0] && pos.X <= rectVertices[0] + 5)
            {
                if (pos.Y >= rectVertices[1] && pos.Y <= rectVertices[3])
                {
                    return EdgeType.LEFT;
                }
            }

            // check if right
            if (pos.X <= rectVertices[2] && pos.X >= rectVertices[2] - 5)
            {
                if (pos.Y >= rectVertices[1] && pos.Y <= rectVertices[3])
                {
                    return EdgeType.RIGHT;
                }
            }

            return EdgeType.NONE;
        }
    }
}
