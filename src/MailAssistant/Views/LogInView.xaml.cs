using Jpp.AddIn.MailAssistant.ViewModels;
using Microsoft.Toolkit.Win32.UI.Controls.Interop.WinRT;
using System.Windows.Controls;

namespace Jpp.AddIn.MailAssistant.Views
{
    /// <summary>
    /// Interaction logic for LogInView.xaml
    /// </summary>
    public partial class LogInView : UserControl
    {
        public LogInView()
        {
            InitializeComponent();
            DataContext = new LogInViewModel(ThisAddIn.Authentication);
            LogInWebView.NavigationStarting += LogInWebView_NavigationStarting;
        }

        private async void LogInWebView_NavigationStarting(object sender, WebViewControlNavigationStartingEventArgs e)
        {
            var vm = (LogInViewModel) DataContext;
            await vm.EvaluateNavigatingUri(e.Uri);
        }
    }
}
