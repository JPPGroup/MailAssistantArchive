namespace Jpp.AddIn.MailAssistant.Abstracts
{
    internal interface IOutlookItem : IWrappedObject
    {
        string Description { get; }
        string Folder { get; }
        int Size { get; }
    } 
}
