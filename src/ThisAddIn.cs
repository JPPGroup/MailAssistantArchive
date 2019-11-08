using Jpp.AddIn.MailAssistant.Backend;
using Jpp.Common.Backend;
using Jpp.Common.Backend.Auth;
using Microsoft.AppCenter;
using Microsoft.AppCenter.Analytics;
using Microsoft.AppCenter.Crashes;
using System;
using System.Collections.Generic;
using Jpp.AddIn.MailAssistant.Wrappers;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Jpp.AddIn.MailAssistant
{
    public partial class ThisAddIn
    {
        #region Instance Variables

        private Outlook.Explorers _explorers;
        private Outlook.Inspectors _inspectors;
        private AppDeploymentCheck _appCheck;

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
            // Start AppCenter
            AppCenter.Start("85ffea91-fbef-4cdf-9e69-ac7c15e3a683", typeof(Analytics), typeof(Crashes));
            //TODO: check if these are to be set before or after start.
            Analytics.SetEnabledAsync(true);
            Crashes.SetEnabledAsync(true);

            using (var user = new AddressEntryWrapper(Application.Session.CurrentUser.AddressEntry))
            {
                AppCenter.SetUserId(user.Address);
            }

            // Initialize variables
            _explorers = Application.Explorers;
            _inspectors = Application.Inspectors;
            _appCheck = new AppDeploymentCheck();

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
            var explorer = Application.ActiveExplorer();
            var window = new OutlookExplorer(explorer);
            Windows.Add(window);
            
            // Hook up event handlers for window
            window.Close += WrappedWindow_Close;
            window.InvalidateControl += WrappedWindow_InvalidateControl;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Unhook event handlers
            _explorers.NewExplorer -= OutlookEvent_Explorers_NewExplorer;
            _inspectors.NewInspector -= OutlookEvent__Inspectors_NewInspector;

            // Dereference objects
            _appCheck.Dispose();
            _appCheck = null;
            _explorers = null;
            _inspectors = null;
            Windows.Clear();
            Windows = null;
            InspectorWindows.Clear();
            InspectorWindows = null;
            Ribbon = null;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonMailAssistantAddIn();
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

        private static void MessageProvider_OnErrorOccurred(object sender, EventArgs e)
        {
            Authentication = new OfficeAddInOAuth(MessageProvider);
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

        #region TestingFolderDetails
        //Code to be re-implemented
        //private const string PR_CREATION_DATE = "http://schemas.microsoft.com/mapi/proptag/0x30070040";
        //
        //internal static void TestingFolderDetails()
        //{
        //    const string baseTestFolderName = "Testing";
        //    string[] testGroupsFolderNames = {"10000-10499","10500-10999","11500-11999","20000-20499","8500-8999","9000-9499","9500-9999"};
        //    //string[] schemeFolderName = {"20000-", "12000 - 12123", "11000 - 11999", "10000 - 10999", "9508 - 9999", "8810-9507" };

        //    var appFolders = _application.GetNamespace(NAMESPACE_TYPE).Folders;
        //    var baseSharedFolder = appFolders.Cast<Outlook.Folder>().FirstOrDefault(folder => folder.Name == BASE_SHARED_FOLDER_NAME);
        //    var baseTestFolder = baseSharedFolder?.Folders.Cast<Outlook.Folder>().FirstOrDefault(folder => folder.Name == baseTestFolderName);
        //    var baseTestGroupFolders = baseTestFolder?.Folders.Cast<Outlook.Folder>().Where(folder => testGroupsFolderNames.Contains(folder.Name));

        //    if (baseTestGroupFolders == null) return;
        //    var baseTTestGroupFolderList = baseTestGroupFolders.ToList();

        //    foreach (var testGroupFolder in baseTTestGroupFolderList)
        //    {
        //        var detailsList = new List<FolderDetails>();

        //        var testGroupFolders = testGroupFolder.Folders;
        //        foreach (Outlook.Folder testFolder in testGroupFolders)
        //        {
        //            var matched = false;

        //            var testFolderCreated = testFolder.PropertyAccessor.GetProperty(PR_CREATION_DATE) as DateTime?;
        //            var testFolderName = testFolder.Name;
        //            var testFolderItems = testFolder.Items.Count;
        //            var charFolderLoc = testFolderName.IndexOf("-", StringComparison.Ordinal);
        //            var testCode = testFolderName.Substring(0, charFolderLoc).Trim();
        //            var testPartCode = new string(testCode.Where(char.IsDigit).ToArray());

        //            if (!matched)
        //            {
        //                detailsList.Add(new FolderDetails
        //                {
        //                    TestFolderName = testFolderName,
        //                    TestFolderCreated = testFolderCreated,
        //                    TestFullCode = testCode,
        //                    TestPartCode = testPartCode,
        //                    CodeMatch = matched,
        //                    TestItemCount = testFolderItems,
        //                });
        //            }

        //            Marshal.ReleaseComObject(testFolder);
        //        }

        //        ExportCsv(detailsList, $"C:\\MailAssistant-AddIn\\{testGroupFolder.Name}.csv");

        //        Marshal.ReleaseComObject(testGroupFolder);
        //    }
        //}

        //private static void ExportCsv<T>(List<T> genericList, string fileName)
        //{
        //    var sb = new StringBuilder();
        //    if (!Directory.Exists(Path.GetDirectoryName(fileName))) Directory.CreateDirectory(Path.GetDirectoryName(fileName));

        //    var header = "";
        //    var info = typeof(T).GetProperties();
        //    if (!File.Exists(fileName))
        //    {
        //        var file = File.Create(fileName);
        //        file.Close();
        //        foreach (var prop in typeof(T).GetProperties())
        //        {
        //            header += prop.Name + ", ";
        //        }
        //        header = header.Substring(0, header.Length - 2);
        //        sb.AppendLine(header);
        //        TextWriter swHeaders = new StreamWriter(fileName, true);
        //        swHeaders.Write(sb.ToString());
        //        swHeaders.Close();


        //        foreach (var obj in genericList)
        //        {
        //            sb = new StringBuilder();
        //            var line = "";
        //            foreach (var prop in info)
        //            {
        //                line += prop.GetValue(obj, null) + ", ";
        //            }
        //            line = line.Substring(0, line.Length - 2);
        //            sb.AppendLine(line);
        //            TextWriter swLines = new StreamWriter(fileName, true);
        //            swLines.Write(sb.ToString());
        //            swLines.Close();
        //        }
        //    }

        //}

        //private class FolderDetails
        //{
        //    private string _testFolderName;
        //    private string _matchFolderName;

        //    public string TestFolderName
        //    {
        //        get => string.IsNullOrEmpty(_testFolderName)? _testFolderName : $"\"{_testFolderName}\""; 
        //        set => _testFolderName = value;
        //    }
        //    public DateTime? TestFolderCreated { get; set; }
        //    public string TestFullCode { get; set; }
        //    public string TestPartCode { get; set; }
        //    public string MatchFolderName
        //    {
        //        get => string.IsNullOrEmpty(_matchFolderName) ? _matchFolderName : $"\"{_matchFolderName}\"";
        //        set => _matchFolderName = value;
        //    }
        //    public string MatchFullCode { get; set; }
        //    public string MatchPartCode { get; set; }
        //    public bool CodeMatch { get; set; }
        //    public int TestItemCount { get; set; }
        //    public int MatchItemCount { get; set; }
        //    public bool CountEqual { get; set; }
        //}

        #endregion
    }
}
