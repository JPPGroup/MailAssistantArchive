using Jpp.Common;
using Jpp.Common.Backend;
using Jpp.Common.Backend.Auth;
using Jpp.Common.Backend.Projects;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace Jpp.AddIn.MailAssistant.Forms
{
    public partial class ProjectListForm : Form
    {
        private readonly BaseOAuthAuthentication _authentication;
        private readonly Projects _projectService;
        private List<ProjectModel> _projectList;
        private string _searchText = null;

        public string SelectedFolder
        {
            get
            {
                if (lstProjects.SelectedItems.Count != 1) return null;
                var item = lstProjects.SelectedItems[0];

                var group = item.SubItems[2].Text;
                var code = item.SubItems[0].Text;
                var name = item.SubItems[1].Text;

                return $"Testing\\{group}\\{code}-{name}";
            }
        }


        public ProjectListForm(BaseOAuthAuthentication authentication, IStorageProvider storage)
        {
            InitializeComponent();

            _authentication = authentication;
            _projectService = new Projects(_authentication, storage);

            ThisAddIn.MessageProvider.ErrorOccurred += MessageProvider_OnErrorOccurred;

            lstProjects.View = View.Details;
            lstProjects.Columns.Add("Code", 100);
            lstProjects.Columns.Add("Name", 400);
            lstProjects.Columns.Add("Subfolder", 200);
        }

        private void MessageProvider_OnErrorOccurred(object sender, EventArgs e)
        {
            Close();
        }

        private void ProjectListForm_Load(object sender, EventArgs e)
        {
            LoadProjectList();
            ActiveControl = txtSearchBox;
        }

        private async void LoadProjectList()
        {
            var baseDateTime = DateTime.Now;

            if (!_authentication.Authenticated) await _authentication.Authenticate();
            Debug.WriteLine($"Authenticate : +{(DateTime.Now - baseDateTime).TotalMilliseconds / 1000}");
            
            baseDateTime = DateTime.Now;
            var result = await _projectService.GetAllProjects();
            Debug.WriteLine($"GetAllProjects : +{(DateTime.Now - baseDateTime).TotalMilliseconds / 1000}");
            
            baseDateTime = DateTime.Now;
            _projectList = result.OrderByDescending(p => p.Code, new ProjectCodeComparer()).ToList();
            Debug.WriteLine($"Stored : +{(DateTime.Now - baseDateTime).TotalMilliseconds / 1000}");
            
            baseDateTime = DateTime.Now;
            PopulateListBox("");
            Debug.WriteLine($"PopulateListBox : +{(DateTime.Now - baseDateTime).TotalMilliseconds / 1000}");
        }

        private void PopulateListBox(string searchText)
        {
            if (_searchText == searchText) return;

            lstProjects.BeginUpdate();
            lstProjects.Items.Clear();

            var projects = !string.IsNullOrEmpty(searchText) 
                ? _projectList.Where(project => project.Code.ToLower().Contains(searchText.ToLower()) || project.Name.ToLower().Contains(searchText.ToLower())) 
                : _projectList;

            lstProjects.Items.AddRange(projects.Select(p => new ListViewItem(new[] { p.Code, p.Name, p.Grouping })).ToArray());
            lstProjects.EndUpdate();

            _searchText = searchText;
        }

        private void TxtSearchBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (sender is TextBox textBox) PopulateListBox(textBox.Text);
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            DialogResult = lstProjects.SelectedItems.Count > 0 ? DialogResult.OK : DialogResult.Cancel;
            Close();
        }

        ~ProjectListForm()
        {
            ThisAddIn.MessageProvider.ErrorOccurred -= MessageProvider_OnErrorOccurred;
        }
    }
}
