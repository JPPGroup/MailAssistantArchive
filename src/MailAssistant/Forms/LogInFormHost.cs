using Jpp.AddIn.MailAssistant.ViewModels;
using Jpp.AddIn.MailAssistant.Views;
using System.Windows.Forms;

namespace Jpp.AddIn.MailAssistant.Forms
{
    public partial class LogInFormHost : Form
    {
        public LogInFormHost(string url)
        {
            InitializeComponent();

            if (!(elementHost.Child is LogInView ctr)) return;

            var vm = (LogInViewModel) ctr.DataContext;
            vm.Host = this;
            vm.Url = url;
        }
    }
}
