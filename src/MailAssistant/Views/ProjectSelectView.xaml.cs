using System.Windows.Controls;
using Jpp.AddIn.MailAssistant.ViewModels;

namespace Jpp.AddIn.MailAssistant.Views
{
    /// <summary>
    /// Interaction logic for ViewProjectSelect.xaml
    /// </summary>
    public partial class ProjectSelectView : UserControl
    {
        public ProjectSelectView()
        {
            InitializeComponent();
            DataContext = new ProjectSelectViewModel(ThisAddIn.Authentication, ThisAddIn.StorageProvider);
        }
    }
}
