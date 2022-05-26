namespace SubmissionCollector.Models
{
    public interface IInventoryItem
    {
        int DisplayOrder { get; set; }
        bool IsSelected { get; set; }
        string Name { get; }
    }
}
