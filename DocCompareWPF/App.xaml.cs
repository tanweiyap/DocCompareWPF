using System.Windows;

namespace DocCompareWPF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LocalLogging.LoggingClass.WriteLog(e.Exception.Message + e.Exception.StackTrace);
            e.Handled = true;
        }
    }
}
