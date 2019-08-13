using System.Collections.Generic;
using Jpp.AddIn.MailAssistant.ViewModels;
using Jpp.AddIn.MailAssistant.Views;
using System.Windows.Forms;

namespace Jpp.AddIn.MailAssistant.Forms
{
    public partial class ProjectSelectFormHost : Form
    {
        public List<string> SelectedFolders { get; set; }

        public ProjectSelectFormHost()
        {
            InitializeComponent();

            if (!(elementHost.Child is ProjectSelectView ctr)) return;

            var vm = (ProjectSelectViewModel)ctr.DataContext;
            vm.Host = this;
        }
    }
}
