﻿using Microsoft.AppCenter.Analytics;
using System.Collections.Generic;
using System.Linq;

namespace Jpp.AddIn.MailAssistant.OutputReports
{
    internal abstract class MoveBaseReport : IMoveReport
    {
        public string DestinationFolder { get; protected set; }
        public List<ItemProperties> Items { get; }
        public RagStatus OverallStatus => GetOverallStatus();
        public int Target { get; set; }
        public int Error { get; set; }
        public int Moved => Items.Count(i => i.Status == ItemStatus.Moved);
        public int Failed => Items.Count(i => i.Status == ItemStatus.Failed);
        public int Duplicate => Items.Count(i => i.Status == ItemStatus.Duplicate);
        public int Skipped => Items.Count(i => i.Status == ItemStatus.Skipped);

        protected MoveBaseReport()
        {
            Items = new List<ItemProperties>();
        }

        public abstract void LogAndShowResults();

        public void AddAndTrackItem(ItemProperties item)
        {
            Items.Add(item);

            Analytics.TrackEvent("Mail item moved", new Dictionary<string, string> {
                { "Description", item.Description },
                { "Status",item.Status.ToString() },
                { "Source", item.Source },
                { "Destination", item.Destination }
            });
        }

        protected void LogAnalytics(string action)
        {
            Analytics.TrackEvent(action, new Dictionary<string, string> {
                { "Folder", DestinationFolder },
                { "Status", OverallStatus.ToString() },
                { "Target", Target.ToString() },
                { "Error", Error.ToString() },
                { "Moved", Moved.ToString() },
                { "Failed", Failed.ToString() },
                { "Skipped", Skipped.ToString() },
                { "Duplicate", Duplicate.ToString() }
            });
        }

        private RagStatus GetOverallStatus()
        {
            if (Failed > 0 || Error > 0) return RagStatus.Red;
            if (Skipped > 0 || Duplicate > 0) return RagStatus.Amber;
            return Moved == Target ? RagStatus.Green : RagStatus.Red;
        }
    }
}
