using Jpp.AddIn.MailAssistant.Properties;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Jpp.AddIn.MailAssistant
{
    [ComVisible(true)]
    public class RibbonMailAssistantAddIn : Office.IRibbonExtensibility
    {
        private readonly Outlook.Application _olApplication;
        
        //Override of constructor to pass a trusted Outlook.Application object
        public RibbonMailAssistantAddIn(Outlook.Application outlookApplication)
        {
            _olApplication = outlookApplication;
        }

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

        public void Ribbon_OnLoad(Office.IRibbonUI ribbonUi)
        {
            ThisAddIn.Ribbon = ribbonUi;
        }

        public Bitmap SendToHub_GetImage(Office.IRibbonControl control) => control.Context switch
            {
                Outlook.AttachmentSelection _ => Resources.SendToHub_Small,
                Outlook.Explorer _ => Resources.SendToHub_Large,
                _ => null
            };
        

        public void SendToHub_OnAction(Office.IRibbonControl control)
        {
            var selection = GetAttachmentSelection(control);
            if (selection == null || selection.Count < 1) return;

            ThisAddIn.CopyAttachments(selection);
        }

        public bool MoveToShared_GetVisible(Office.IRibbonControl control)
        {
            var selection = GetItemSelection(control);
            if (selection == null || selection.Count < 1) return false;

            for (var i = 1; i <= selection.Count; i++)
            {
                if (!(selection[i] is Outlook.MailItem oMail)) return false;
                if (!oMail.Sent) return false;
            }

            return true;
        }

        public void MoveToShared_OnAction(Office.IRibbonControl control)
        {
            var selection = GetItemSelection(control);
            if (selection == null || selection.Count < 1) return;

            ThisAddIn.MoveMail(selection);
        }

        public Bitmap MoveToShared_GetImage(Office.IRibbonControl control) => control.Context switch
            { 
                Outlook.Selection _ => Resources.MoveToShared_Small, 
                Outlook.Explorer _ => Resources.MoveToShared_Large, 
                _ => null
            };

        public void MoveFolderToSharedFolder_OnAction(Office.IRibbonControl control)
        {
            var folder = control.Context as Outlook.Folder;
            ThisAddIn.MoveFolderContents(folder);
        }

        public Bitmap MoveFolderToSharedFolder_GetImage(Office.IRibbonControl control)
        {
            return Resources.MoveFolderToShared_Small;
        }

        public void NewFolder_OnAction(Office.IRibbonControl control)
        {
            ThisAddIn.NewFolder();
        }

        public Bitmap NewFolder_GetImage(Office.IRibbonControl control)
        {
            return Resources.NewSharedFolder_Large;
        }

        #endregion

        #region Helpers

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
                _ => null
            };

        private Outlook.Selection GetItemSelection(Office.IRibbonControl control) => control.Context switch
            {
                Outlook.Selection context => context,
                Outlook.Explorer explorer => explorer.Selection,
                _ => null
            };

        #endregion
    }
}
