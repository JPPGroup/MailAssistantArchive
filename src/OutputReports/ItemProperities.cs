namespace Jpp.AddIn.MailAssistant.OutputReports
{
    public class ItemProperties
    {
        public string Description { get; }
        public string Source { get; }
        public string Destination { get; }
        public int Size { get; }
        public ItemStatus Status { get; }

        public ItemProperties(string description, string source, string destination, int size, ItemStatus status)
        {
            Description = description;
            Source = source;
            Destination = destination;
            Size = size;
            Status = status;
        }
    }
}
