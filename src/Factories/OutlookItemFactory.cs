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
 * https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/outlook-item-objects
 *
 * Outlook.AppointmentItem - Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder.
 * Outlook.ContactItem - Represents a contact in a Contacts folder.
 * Outlook.DistListItem - Represents a distribution list in a Contacts folder.
 * Outlook.DocumentItem - Represents any document other than a Microsoft Outlook item as an item in an Outlook folder.
 * Outlook.JournalItem - Represents a journal entry in a Journal folder.
 * Outlook.NoteItem - Represents a note in a Notes folder.
 * Outlook.PostItem - Represents a post in a public folder that others may browse.
 * Outlook.RemoteItem - Represents a remote item in an Inbox folder.
 * Outlook.SharingItem - Represents a sharing message in an Inbox folder.
 * Outlook.StorageItem - A message object in MAPI that is always saved as a hidden item in the parent folder and stores private data for Outlook solutions.
 * Outlook.TaskItem - Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.
 * Outlook.TaskRequestAcceptItem - Represents a response to a TaskRequestItem sent by the initiating user. 
 * Outlook.TaskRequestDeclineItem - Represents a response to a TaskRequestItem sent by the initiating user.
 * Outlook.TaskRequestItem - Represents a change to the recipient's Tasks list initiated by another party or as a result of a group tasking.
 * Outlook.TaskRequestUpdateItem - Represents a response to a TaskRequestItem sent by the initiating user.
 */

