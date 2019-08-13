using Jpp.Common.Backend;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Jpp.AddIn.MailAssistant.Backend
{
    internal class MessageProvider : IMessageProvider
    {
        public Task ShowStorageAccessPermissionWarning()
        {
            throw new NotImplementedException();
        }

        public async Task ShowCriticalError(string message)
        {
            await ShowError(message);
        }

        public async Task ShowError(string message)
        {
            await Task.Run(() => MessageBox.Show(message));
        }
    }
}
