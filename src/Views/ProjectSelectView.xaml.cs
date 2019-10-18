using System;
using System.Windows.Forms;
using Jpp.AddIn.MailAssistant.ViewModels;
using UserControl = System.Windows.Controls.UserControl;

namespace Jpp.AddIn.MailAssistant.Views
{
    /// <summary>
    /// Interaction logic for ViewProjectSelect.xaml
    /// </summary>
    public partial class ProjectSelectView : UserControl
    {
        public Form FormsWindow { get; set; }

        public ProjectSelectView()
        {
            InitializeComponent();
            DataContext = new ProjectSelectViewModel(ThisAddIn.Authentication, ThisAddIn.StorageProvider);
            ThisAddIn.MessageProvider.ErrorOccurred += MessageProvider_OnErrorOccurred;
        }

        private void MessageProvider_OnErrorOccurred(object sender, EventArgs e)
        {
            FormsWindow.DialogResult = DialogResult.Cancel;
            FormsWindow.Close();
        }

        ~ProjectSelectView()
        {
            ThisAddIn.MessageProvider.ErrorOccurred -= MessageProvider_OnErrorOccurred;
        }
    }
}
