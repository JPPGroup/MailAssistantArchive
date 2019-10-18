using Jpp.AddIn.MailAssistant.Forms;
using Jpp.Common;
using Jpp.Common.Backend;
using Jpp.Common.Backend.Auth;
using Jpp.Common.Backend.Projects;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;

namespace Jpp.AddIn.MailAssistant.ViewModels
{
    internal class ProjectSelectViewModel : BaseNotify
    {
        private readonly BaseOAuthAuthentication _authentication;
        private readonly Projects _projectService;
        private ICommand _cancelCommand;
        private ICommand _okCommand;
        private string _searchText;
        private ListCollectionView _projectsView;
        private ProjectModel _selectedProject;

        public ProjectModel SelectedProject
        {
            get => _selectedProject;
            set
            {
                SetField(ref _selectedProject, value, nameof(SelectedProject));
                Host.SelectedFolders = value is null ? null : new List<string> { $"Testing\\{_selectedProject.Grouping}\\{_selectedProject.Code}-{_selectedProject.Name}" };
            }
        }
        public ICollectionView ProjectsView
        {
            get => _projectsView;
            set => SetField(ref _projectsView, (ListCollectionView)value, nameof(ProjectsView));
        }
        public string SearchText {
            get => _searchText;
            set
            {
                SetField(ref _searchText, value, nameof(SearchText));
                OnPropertyChanged(nameof(SearchBackgroundVisible));

                if (string.IsNullOrEmpty(value))
                    ProjectsView.Filter = null;
                else
                    ProjectsView.Filter = o => 
                            ((ProjectModel)o).Code.ToLower().Contains(value.ToLower()) || 
                            ((ProjectModel)o).Name.ToLower().Contains(value.ToLower());
            }
        }
        public Visibility SearchBackgroundVisible => string.IsNullOrWhiteSpace(SearchText) ? Visibility.Visible : Visibility.Hidden;
        public ProjectSelectFormHost Host { get; set; }
        public ICommand CancelCommand => _cancelCommand ??= new DelegateCommand(DoCancel);
        public ICommand OkCommand => _okCommand ??= new DelegateCommand(DoOk);

        public ProjectSelectViewModel(BaseOAuthAuthentication authentication, IStorageProvider storage)
        {
            _authentication = authentication;
            _projectService = new Projects(_authentication, storage);

            LoadProjectList();
        }

        private void DoCancel()
        {
            Host.DialogResult = DialogResult.Cancel;
            Host.Close();
        }

        private void DoOk()
        {
            Host.DialogResult = DialogResult.OK;
            Host.Close();
        }

        private async void LoadProjectList()
        {
            if (!_authentication.Authenticated) await _authentication.Authenticate();

            var result = await _projectService.GetAllProjects();
            var list = result.OrderByDescending(p => p.Code, new ProjectCodeComparer()).ToList();
            
            ProjectsView = new ListCollectionView(list);
            SelectedProject = null;
        }
    }
}
