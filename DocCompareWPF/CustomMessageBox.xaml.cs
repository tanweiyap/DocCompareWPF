using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Documents;

namespace DocCompareWPF
{
    /// <summary>
    /// Interaktionslogik für CustomMessageBox.xaml
    /// </summary>
    public partial class CustomMessageBox : Window
    {
        public CustomMessageBox()
        {
            InitializeComponent();
        }

        public void Setup(string Title, string Message, string ButtonText = "Okay", string Button1Text = "")
        {
            WindowTitle.Content = Title;
            Button2.Content = ButtonText;
            ParseText(Message);
            if (Button1Text.Length != 0)
            {
                Button1.Content = Button1Text;
                Button1.Visibility = Visibility.Visible;
            }
        }

        private void ParseText(string text)
        {
            if(text.Contains("www.hopie.tech"))
            {
                string[] splitText = text.Split("www.hopie.tech");
                Span spanItem = new Span();

                spanItem.Inlines.Add(splitText[0]);

                Hyperlink hyperLink = new Hyperlink()
                {
                    NavigateUri = new Uri("https://www.hopie.tech")
                };
                hyperLink.Inlines.Add("www.hopie.tech");
                hyperLink.RequestNavigate += Hyperlink_RequestNavigate;

                spanItem.Inlines.Add(hyperLink);
                spanItem.Inlines.Add(splitText[^1]);
                MessageTextBlock.Inlines.Add(spanItem);
            }
            else
            {

                MessageTextBlock.Text = text;
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

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}