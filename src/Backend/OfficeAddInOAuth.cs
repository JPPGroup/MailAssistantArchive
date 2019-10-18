using Jpp.AddIn.MailAssistant.Forms;
using Jpp.Common.Backend;
using Jpp.Common.Backend.Auth;

namespace Jpp.AddIn.MailAssistant.Backend
{
    internal class OfficeAddInOAuth : BaseOAuthAuthentication
    {
        public OfficeAddInOAuth(IMessageProvider messenger) : base(messenger) { }

        public override void AuthenticationPrompt()
        {
            using var form = new LogInFormHost(GetAuthenticationURL());
            form.ShowDialog();
        }

        public override void ExpirePrompt()
        {
            using var form = new LogInFormHost(GetExpireURL());
            form.ShowDialog();
        }
    }
}