using Jpp.AddIn.MailAssistant.Factories;
using Jpp.AddIn.MailAssistant.Properties;
using Jpp.AddIn.MailAssistant.Wrappers;
using Microsoft.AppCenter.Crashes;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Jpp.AddIn.MailAssistant
{
    [ComVisible(true)]
    public class RibbonMailAssistantAddIn : Office.IRibbonExtensibility
    {
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            string customUi;

            //Return the appropriate Ribbon XML for ribbonId
            switch (ribbonId)
            {
                case "Microsoft.Outlook.Explorer":
                    customUi = GetResourceText("Jpp.AddIn.MailAssistant.Ribbons.Explorer.xml");
                    return customUi;
                case "Microsoft.Outlook.Mail.Read":
                    customUi = GetResourceText("Jpp.AddIn.MailAssistant.Ribbons.ReadMail.xml");
                    return customUi;
                default:
                    return string.Empty;
            }
        }

        #endregion

        #region Ribbon Callbacks

        public void OnLoad_Ribbon(Office.IRibbonUI ribbonUi)
        {
            ThisAddIn.Ribbon = ribbonUi;
        }

        public Bitmap GetImage_SendToHub(Office.IRibbonControl control) => control.Context switch
            {
                Outlook.AttachmentSelection _ => Resources.SendToHub_Small,
                Outlook.Explorer _ => Resources.SendToHub_Large,
                _ => null
            };
        
        public void OnAction_SendToHub(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetAttachmentSelection(control);
                CopyAttachments(selection);
            }
            catch (Exception e)
            {
                Crashes.TrackError(e);
                //TODO: need to info user. Cannot rethrow as will be swallowed up by Outlook.
            }
        }

        public bool GetVisible_MoveToShared(Office.IRibbonControl control)
        {
            var selection = GetItemSelection(control);
            return selection != null && selection.Count >= 1;
        }

        public void OnAction_MoveToShared(Office.IRibbonControl control)
        {
            try
            {
                var selection = GetItemSelection(control);
                MoveSelection(selection);
            }
            catch (Exception e)
            {
                Crashes.TrackError(e);
                //TODO: need to info user. Cannot rethrow as will be swallowed up by Outlook.
            }
        }

        public Bitmap GetImage_MoveToShared(Office.IRibbonControl control) => control.Context switch
            { 
                Outlook.Selection _ => Resources.MoveToShared_Small, 
                Outlook.Explorer _ => Resources.MoveToShared_Large, 
                _ => null
            };

        public void OnAction_MoveFolderToSharedFolder(Office.IRibbonControl control)
        {
            try
            {
                var folder = control.Context as Outlook.Folder;
                MoveFolder(folder);
            }
            catch (Exception e)
            {
                Crashes.TrackError(e);
                //TODO: need to info user. Cannot rethrow as will be swallowed up by Outlook.
            }
        }

        public Bitmap GetImage_MoveFolderToSharedFolder(Office.IRibbonControl control) //IRibbonControl need for callback signature but unused
        {
            return Resources.MoveFolderToShared_Small;
        }

        public void OnAction_NewFolder(Office.IRibbonControl control) //IRibbonControl need for callback signature but unused
        {
            try
            {
                NewSharedFolder();
            }
            catch (Exception e)
            {
                Crashes.TrackError(e);
                //TODO: need to info user. Cannot rethrow as will be swallowed up by Outlook.
            }
        }

        public Bitmap GetImage_NewFolder(Office.IRibbonControl control) //IRibbonControl need for callback signature but unused
        {
            return Resources.NewSharedFolder_Large;
        }

        public void OnAction_TestDetails(Office.IRibbonControl control) //IRibbonControl need for callback signature but unused
        {
            try
            {
                throw new NotImplementedException(); //TODO: re-implement this based on wrappers
                //ThisAddIn.TestingFolderDetails();
            }
            catch (Exception e)
            {
                Crashes.TrackError(e);
                //TODO: need to info user. Cannot rethrow as will be swallowed up by Outlook.
            }
        }

        public Bitmap GetImage_TestDetails(Office.IRibbonControl control) //IRibbonControl need for callback signature but unused
        {
            return Resources.TestDetails__Large;
        }

        #endregion

        #region Helpers

        private void MoveSelection(Outlook.Selection selection)
        {
            using (var wrappedSelection = new SelectionWrapper(selection))
            using (var wrappedSharedFolder = OutlookFolderFactory.GetOrCreateSharedFolder(Globals.ThisAddIn.Application))
            {
                wrappedSharedFolder.MoveIntoFolder(wrappedSelection);
            }
        }

        private void MoveFolder(Outlook.Folder folder)
        {
            using (var wrappedFolder = new FolderWrapper(folder))
            using (var wrappedSharedFolder = OutlookFolderFactory.GetOrCreateSharedFolder(Globals.ThisAddIn.Application))
            {
                wrappedSharedFolder.MoveIntoFolder(wrappedFolder);
            }
        }

        private void NewSharedFolder()
        {
            using var folder = OutlookFolderFactory.GetOrCreateSharedFolder(Globals.ThisAddIn.Application);
            {
                var stringBuilder = new StringBuilder();
                stringBuilder.Append($"Successfully created folder: \n{folder.Name}.");

                MessageBox.Show(stringBuilder.ToString(), @"Mail Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static void CopyAttachments(Outlook.AttachmentSelection selection)
        {
            MessageBox.Show(@"Not implemented", @"Mail Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) != 0) continue;

                var stream = asm.GetManifestResourceStream(name);
                if (stream == null) continue;

                using var resourceReader = new StreamReader(stream);
                return resourceReader.ReadToEnd();
            }
            return null;
        }

        private Outlook.AttachmentSelection GetAttachmentSelection(Office.IRibbonControl control) => control.Context switch
            {
                Outlook.AttachmentSelection context => context,
                Outlook.Explorer explorer => explorer.AttachmentSelection,
                _ => throw new NotImplementedException()
            };

        private Outlook.Selection GetItemSelection(Office.IRibbonControl control) => control.Context switch
            {
                Outlook.Selection context => context,
                Outlook.Explorer explorer => explorer.Selection,
                _ => throw new NotImplementedException()
            };

        #endregion
    }
}
