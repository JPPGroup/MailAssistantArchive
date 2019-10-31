using System.Collections.Generic;

namespace Jpp.AddIn.MailAssistant.OutputReports
{
    public interface IMoveReport
    {
        string DestinationFolder { get; }
        List<ItemProperties> Items { get; }
        RagStatus OverallStatus { get; }
        int Target { get; }
        int Error { get; }
        int Moved { get; }
        int Failed { get; }
        int Duplicate { get; }
        int Skipped { get; }

        void LogAndShowResults();
        void AddAndTrackItem(ItemProperties item);
    }
}
