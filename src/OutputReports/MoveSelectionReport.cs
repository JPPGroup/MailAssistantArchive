using Jpp.AddIn.MailAssistant.Forms;
using Jpp.AddIn.MailAssistant.Wrappers;

namespace Jpp.AddIn.MailAssistant.OutputReports
{
    internal class MoveSelectionReport : MoveBaseReport
    {
        public FolderProperties Before { get; }
        public FolderProperties After { get; }
        
        public MoveSelectionReport(FolderWrapper destinationFolder, SelectionWrapper selection)
        {
            DestinationFolder = destinationFolder.Name;
            Target = selection.Count;

            Before = new FolderProperties
            {
                Name = destinationFolder.Name,
                Size = destinationFolder.Size,
                NoOfFolders = destinationFolder.Folders.Count,
                NoOfItems = destinationFolder.Items.Count
            };

            After = new FolderProperties { Name = destinationFolder.Name };
        }

        public override void LogAndShowResults()
        {
            LogAnalytics("Selection move complete");

            using var frmResult = new MoveReportForm(this);
            frmResult.ShowDialog();
        }

        public void SetAfterDetails(FolderWrapper destinationFolder)
        {
            After.Size = destinationFolder.Size;
            After.NoOfFolders = destinationFolder.Folders.Count;
            After.NoOfItems = destinationFolder.Items.Count;
        }
    }
}
