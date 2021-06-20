using DocCompareWPF.Classes;
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

        private enum EditMode
        {
            SELECT,
            CREATE,
        };

        public NoCompareZoneWindow()
        {
            InitializeComponent();

            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            SelectButtonBackground.Background = brush;
            CreateRectBackground.Background = Brushes.Transparent;
            mode = EditMode.SELECT;
            rects = new List<List<double>>();
        }

        public void SetupWindow(Document doc)
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
        }

        private void CreateRectButton_Click(object sender, RoutedEventArgs e)
        {
            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            SelectButtonBackground.Background = Brushes.Transparent;
            CreateRectBackground.Background = brush;
            mode = EditMode.CREATE;
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            Brush brush = FindResource("SecondaryAccentBrush") as Brush;
            SelectButtonBackground.Background = brush;
            CreateRectBackground.Background = Brushes.Transparent;
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

            if(pageIndex == 0)
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

            if(pageIndex == fi.Length -1)
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
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void PageCanvas_MouseEnter(object sender, MouseEventArgs e)
        {
            if(mode == EditMode.CREATE)
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
            if(mode == EditMode.CREATE)
            {
                isMouseDown = true;
                startPoint = e.GetPosition(PageCanvas);

                System.Windows.Shapes.Rectangle newRect = new System.Windows.Shapes.Rectangle();
                newRect.Stroke = Brushes.Black;

                Brush brush = FindResource("WarningBackgroundBrush") as Brush;

                newRect.Fill = brush;
                newRect.Opacity = 0.5;
                newRect.Width = 1;
                newRect.Height = 1;

                Canvas.SetLeft(newRect, startPoint.X);
                Canvas.SetTop(newRect, startPoint.Y);
                PageCanvas.Children.Add(newRect);
            }
        }

        private void PageCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (mode == EditMode.CREATE && isMouseDown == true)
            {
                endPoint = e.GetPosition(PageCanvas);
                System.Windows.Shapes.Rectangle newRect = new System.Windows.Shapes.Rectangle();
                newRect.Stroke = Brushes.Black;

                Brush brush = FindResource("WarningBackgroundBrush") as Brush;

                newRect.Fill = brush;
                newRect.Opacity = 0.5;
                newRect.Width = Math.Abs(endPoint.X - startPoint.X);
                newRect.Height = Math.Abs(endPoint.Y - startPoint.Y);

                Canvas.SetLeft(newRect, Math.Min(startPoint.X, endPoint.X));
                Canvas.SetTop(newRect, Math.Min(startPoint.Y, endPoint.Y));
                PageCanvas.Children.RemoveAt(PageCanvas.Children.Count - 1);
                PageCanvas.Children.Add(newRect);

                if(System.Windows.Input.Mouse.LeftButton == MouseButtonState.Released)
                {
                    isMouseDown = false;
                    List<double> values = new List<double>();
                    values.Add(Math.Min(startPoint.X, endPoint.X));
                    values.Add(Math.Min(startPoint.Y, endPoint.Y));
                    values.Add(Math.Max(startPoint.X, endPoint.X));
                    values.Add(Math.Max(startPoint.Y, endPoint.Y));
                    rects.Add(values);
                }
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape && mode == EditMode.CREATE)
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
        }
    }
}
