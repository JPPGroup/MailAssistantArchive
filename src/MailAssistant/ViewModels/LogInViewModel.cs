using Jpp.AddIn.MailAssistant.Forms;
using Jpp.Common;
using Jpp.Common.Backend.Auth;
using System;
using System.Threading.Tasks;

namespace Jpp.AddIn.MailAssistant.ViewModels
{
    internal class LogInViewModel : BaseNotify
    {
        private readonly BaseOAuthAuthentication _authentication;
        private string _url;

        public LogInFormHost Host { get; set; }
        public string Url
        {
            get => _url;
            set => SetField(ref _url, value, nameof(Url));
        }

        public LogInViewModel(BaseOAuthAuthentication authentication)
        {
            _authentication = authentication;
        }

        public async Task EvaluateNavigatingUri(Uri uri)
        {        
            if (await _authentication.EvaluateURL(uri)) Host.Close();
        }

        public void Logout()
        {
            throw new NotImplementedException();
        }

    }
}
