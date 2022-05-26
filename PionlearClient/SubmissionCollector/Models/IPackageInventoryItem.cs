using System.Collections.Generic;

namespace SubmissionCollector.Models
{
    public interface IPackageInventoryItem : IInventoryItem
    {
        string CedentIdAndName { get; }
        IEnumerable<IInventoryItem> SegmentViews { get; }
    }
}