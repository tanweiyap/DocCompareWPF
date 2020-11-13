using System.Windows;

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
            MessageTextBlock.Text = Message;
            Button2.Content = ButtonText;
            if (Button1Text.Length != 0)
            {
                Button1.Content = Button1Text;
                Button1.Visibility = Visibility.Visible;
            }
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            Close();
            return;
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            Close();
            return;
        }
    }
}