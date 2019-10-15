using Jpp.AddIn.MailAssistant.Backend;
using Jpp.Common.Backend;
using Jpp.Common.Backend.Auth;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Jpp.AddIn.MailAssistant.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Jpp.AddIn.MailAssistant
{
    public partial class ThisAddIn
    {
        private const string BASE_SHARED_FOLDER_NAME = "JPP_Shared";
        private const string NAMESPACE_TYPE = "MAPI";
        private const int SEARCH_WINDOW_MINUTES = 2;
        private const string PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E";
        private const string SEARCH_DATE_FORMAT = "dd MMMM yyyy h:mm tt";

        #region Instance Variables

        private Outlook.Explorers _explorers;
        private Outlook.Inspectors _inspectors;
        

        private static Outlook.Application _application;

        internal static List<OutlookExplorer> Windows;  // List of tracked explorer windows  
        internal static List<OutlookInspector> InspectorWindows; // List of tracked inspector windows         
        internal static Office.IRibbonUI Ribbon; // Ribbon UI reference
        internal static BaseOAuthAuthentication Authentication;
        internal static IStorageProvider StorageProvider;
        internal static MessageProvider MessageProvider;

        #endregion

        #region VSTO Startup and Shutdown methods

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Initialize variables
            _application = Application;
            _explorers = _application.Explorers;
            _inspectors = _application.Inspectors;

            MessageProvider = new MessageProvider();
            Windows = new List<OutlookExplorer>();
            InspectorWindows = new List<OutlookInspector>();
            StorageProvider = new StorageProvider();
            Authentication = new OfficeAddInOAuth(MessageProvider);

            // Wire up event handlers to handle multiple Explorer windows
            _explorers.NewExplorer += OutlookEvent_Explorers_NewExplorer;

            // Wire up event handlers to handle multiple Inspector windows
            _inspectors.NewInspector += OutlookEvent__Inspectors_NewInspector;

            MessageProvider.ErrorOccurred += MessageProvider_OnErrorOccurred;
            
            // Add the ActiveExplorer to Windows
            var explorer = _application.ActiveExplorer();
            var window = new OutlookExplorer(explorer);
            Windows.Add(window);
            
            // Hook up event handlers for window
            window.Close += WrappedWindow_Close;
            window.InvalidateControl += WrappedWindow_InvalidateControl;
        }

        private  void MessageProvider_OnErrorOccurred(object sender, EventArgs e)
        {
            Authentication = new OfficeAddInOAuth(MessageProvider);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Unhook event handlers
            _explorers.NewExplorer -= OutlookEvent_Explorers_NewExplorer;
            _inspectors.NewInspector -= OutlookEvent__Inspectors_NewInspector;

            // Dereference objects
            _explorers = null;
            _inspectors = null;
            Windows.Clear();
            Windows = null;
            InspectorWindows.Clear();
            InspectorWindows = null;
            Ribbon = null;
            _application = null;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonMailAssistantAddIn(_application);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Looks up the window wrapper for a given window object
        /// </summary>
        /// <param name="window">An outlook explorer window</param>
        /// <returns></returns>
        internal static OutlookExplorer FindOutlookExplorer(object window)
        {
            foreach (var explorer in Windows)
            {
                if (explorer.Window == window) return explorer;
            }

            return null;
        }

        /// <summary>
        /// Looks up the window wrapper for a given window object
        /// </summary>
        /// <param name="window">An outlook inspector window</param>
        /// <returns></returns>
        internal static OutlookInspector FindOutlookInspector(object window)
        {
            foreach (var inspector in InspectorWindows)
            {
                if (inspector.Window == window) return inspector;
            }

            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="selection"></param>
        internal static void MoveMail(Outlook.Selection selection)
        {
            if (selection == null || selection.Count < 1) return;

            using var form = new ProjectSelectFormHost();
            var result = form.ShowDialog();

            if (result != DialogResult.OK) return;

            var folder = GetSharedFolder(form.SelectedFolders[0]);
            if(folder == null) throw new ArgumentNullException(nameof(folder), @"No shared folder.");

            var duplicates = new List<string>();
            foreach (var item in selection)
            {
                if (item is Outlook.MailItem mail)
                {
                    if (IsDuplicateInFolder(folder, mail))
                    {
                        duplicates.Add(mail.Subject);
                        continue;
                    }

                    mail.Move(folder);
                    //TODO: Log analytics  
                }
            }

            if (!duplicates.Any()) return;
            
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("The following items where already present in the folder: \n");
            foreach (var item in duplicates)
            {
                stringBuilder.AppendLine($"\n{item}");
            }

            MessageBox.Show(stringBuilder.ToString(), @"Mail Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="delete"></param>
        internal static void MoveFolderContents(Outlook.Folder folder, bool delete)
        {
            if (folder == null) return;
            using var form = new ProjectSelectFormHost();
            var result = form.ShowDialog();

            if (result != DialogResult.OK) return;

            var sharedFolder= GetSharedFolder(form.SelectedFolders[0]);
            if (sharedFolder == null) throw new ArgumentNullException(nameof(sharedFolder), @"No shared folder.");
            foreach (var item in folder.Items)
            {
                if (item is Outlook.MailItem mail)
                {
                    mail.Move(sharedFolder);
                }
            }

            if (delete) folder.Delete();
        }

        /// <summary>
        /// 
        /// </summary>
        internal static void NewFolder()
        {
            using var form = new ProjectSelectFormHost();
            var result = form.ShowDialog();

            if (result != DialogResult.OK) return;

            var _ = GetSharedFolder(form.SelectedFolders[0]);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="selection"></param>
        internal static void CopyAttachments(Outlook.AttachmentSelection selection)
        {
            MessageBox.Show(@"Not implemented", @"Mail Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region Event Handlers

        private static void OutlookEvent_Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            // Check to see if this is a new window we don't already track
            var existingWindow = FindOutlookExplorer(explorer);

            // If the collection has a window for this Explorer then return, otherwise we should add it
            if (existingWindow != null) return;

            var window = new OutlookExplorer(explorer);
            window.Close += WrappedWindow_Close;
            window.InvalidateControl += WrappedWindow_InvalidateControl;
            Windows.Add(window);
        }

        private static void OutlookEvent__Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            Ribbon.Invalidate();

            // Check to see if this is a new window we don't already track
            var existingInspector = FindOutlookInspector(inspector);
            
            // If the collection has a window for this Inspector then return, otherwise we should add it
            if (existingInspector != null) return;

            var window = new OutlookInspector(inspector);
            window.Close += WrappedInspectorWindow_Close;
            window.InvalidateControl += WrappedInspectorWindow_InvalidateControl;
            InspectorWindows.Add(window);
        }

        private static void WrappedInspectorWindow_InvalidateControl(object sender, OutlookInspector.InvalidateEventArgs e)
        {
            Ribbon?.InvalidateControl(e.ControlId);
        }

        private static void WrappedInspectorWindow_Close(object sender, EventArgs e)
        {
            var window = (OutlookInspector)sender;
            window.Close -= WrappedInspectorWindow_Close;
            InspectorWindows.Remove(window);
        }

        private static void WrappedWindow_InvalidateControl(object sender, OutlookExplorer.InvalidateEventArgs e)
        {
            Ribbon?.InvalidateControl(e.ControlId);
        }

        private static void WrappedWindow_Close(object sender, EventArgs e)
        {
            var window = (OutlookExplorer)sender;
            window.Close -= WrappedWindow_Close;
            Windows.Remove(window);
        }

        #endregion

        #region Helpers

        private static Outlook.Folder GetSharedFolder(string folderName)
        {
            if (string.IsNullOrEmpty(folderName)) throw new ArgumentNullException(nameof(folderName), @"Folder name not set.");

            var sharedFolder = GetFolder(_application.GetNamespace(NAMESPACE_TYPE).Folders, BASE_SHARED_FOLDER_NAME);
            if (sharedFolder == null) throw new ArgumentNullException(nameof(sharedFolder), @"Base shared folder not set.");

            var arrFolders = folderName.Split('\\');
            var folder = sharedFolder;

            for (var i = 0; i <= arrFolders.GetUpperBound(0); i++)
            {
                var colFolders = folder.Folders;
                var nextFolder = GetFolder(colFolders, arrFolders[i]) ?? CreateFolder(folder, arrFolders[i]);

                folder = nextFolder;
            }

            return folder;
        }

        private static Outlook.Folder GetFolder(Outlook.Folders folders, string folderName)
        {
            return folders.Cast<Outlook.Folder>().FirstOrDefault(folder => folder.Name == folderName);
        }

        private static Outlook.Folder CreateFolder(Outlook.Folder rootFolder, string folderName)
        {
            return (Outlook.Folder)rootFolder.Folders.Add(folderName, Outlook.OlDefaultFolders.olFolderInbox);
        }

        private static bool IsDuplicateInFolder(Outlook.Folder folder, Outlook.MailItem mail)
        {
            Outlook.Items folderItems = null;
            Outlook.Items resultItems = null;
            object item = null;

            try
            {
                var dateFrom = mail.SentOn.AddMinutes(-SEARCH_WINDOW_MINUTES).ToString(SEARCH_DATE_FORMAT);
                var dateTo = mail.SentOn.AddMinutes(SEARCH_WINDOW_MINUTES).ToString(SEARCH_DATE_FORMAT);
                var mailId = GetMessageId(mail);
                var restrictCriteria = $"[SentOn] >= '{dateFrom}' And [SentOn] <= '{dateTo}'";

                folderItems = folder.Items;
                resultItems = folderItems.Restrict(restrictCriteria);
                item = resultItems.GetFirst();

                while (item != null)
                {
                    if (item is Outlook.MailItem mailItem)
                    {
                        if (GetMessageId(mailItem) == mailId) return true;
                        if (mailItem.ReceivedTime == mail.ReceivedTime && mailItem.Subject == mail.Subject) return true;
                    }

                    Marshal.ReleaseComObject(item);
                    item = resultItems.GetNext();
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (item != null) Marshal.ReleaseComObject(item);
                if (folderItems != null) Marshal.ReleaseComObject(folderItems);
                if (resultItems != null) Marshal.ReleaseComObject(resultItems);
            }
        }

        private static string GetMessageId(Outlook.MailItem mail)
        {
            try
            {
                return mail.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) as string;
            }
            catch (Exception)
            {
                return "";
            }
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
