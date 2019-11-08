using Jpp.AddIn.MailAssistant.Forms;
using Jpp.AddIn.MailAssistant.Wrappers;

namespace Jpp.AddIn.MailAssistant.OutputReports
{
    internal class MoveFolderReport : MoveBaseReport
    {
        public FolderProperties SourceBefore { get; }
        public FolderProperties DestinationBefore { get; }

        public FolderProperties SourceAfter { get; }
        public FolderProperties DestinationAfter { get; }

        public MoveFolderReport(FolderWrapper source, FolderWrapper destination)
        {
            Target = source.Items.Count;
            DestinationFolder = destination.Name;

            SourceBefore = new FolderProperties
            {
                Name = source.Name,
                NoOfItems = source.Items.Count,
                NoOfFolders = source.Folders.Count,
                Size = source.Size
            };

            DestinationBefore = new FolderProperties
            {
                Name = destination.Name,
                NoOfItems = destination.Items.Count,
                NoOfFolders = destination.Folders.Count,
                Size = destination.Size
            };

            SourceAfter = new FolderProperties { Name = source.Name };
            DestinationAfter = new FolderProperties { Name = destination.Name };
        }

        public override void LogAndShowResults()
        {
            LogAnalytics("Folder move complete");

            using var frmResult = new MoveReportForm(this);
            frmResult.ShowDialog();
        }

        public void SetAfterDetails(FolderWrapper source, FolderWrapper destination)
        {
            SourceAfter.Size = source.Size;
            SourceAfter.NoOfFolders = source.Folders.Count;
            SourceAfter.NoOfItems = source.Items.Count;

            DestinationAfter.Size = destination.Size;
            DestinationAfter.NoOfFolders = destination.Folders.Count;
            DestinationAfter.NoOfItems = destination.Items.Count;
        }
    }
}
