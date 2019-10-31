using Jpp.AddIn.MailAssistant.Abstracts;
using Jpp.AddIn.MailAssistant.Exceptions;
using Jpp.AddIn.MailAssistant.Wrappers;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Jpp.AddIn.MailAssistant.Factories
{
    internal static class OutlookItemFactory
    {
        public static IOutlookItem Create(dynamic item)
        {
            switch (item)
            {
                case Outlook.MailItem mailItem:
                    return new MailItemWrapper(mailItem);
                case Outlook.MeetingItem meetingItem:
                    return new MeetingItemWrapper(meetingItem);
                case Outlook.ReportItem reportItem:
                    return new ReportItemWrapper(reportItem);
                default:
                    throw new OutlookItemFactoryException(@"Outlook item type not handled", nameof(item));
            }

        }
    }
}
/*
 * Possible types to add later
 *
 * Outlook.AppointmentItem appointmentItem
 * Outlook.ContactItem contactItem
 * Outlook.DistListItem distListItem
 * Outlook.DocumentItem documentItem
 * Outlook.JournalItem journalItem
 * Outlook.NoteItem noteItem
 * Outlook.PostItem postItem
 * Outlook.RemoteItem remoteItem
 * Outlook.SharingItem sharingItem
 * Outlook.StorageItem storageItem
 * Outlook.TaskItem taskItem
 * Outlook.TaskRequestAcceptItem taskRequestAcceptItem 
 * Outlook.TaskRequestDeclineItem taskRequestDeclineItem
 * Outlook.TaskRequestItem taskRequestItem
 * Outlook.TaskRequestUpdateItem taskRequestUpdateItem
 */

