using System.Collections.Generic;

namespace SubmissionCollector.Models
{
    public interface ISegmentInventoryItem : IInventoryItem
    {
        IEnumerable<IInventoryItem> SublineViews { get; }
        bool IsUmbrella { get; }
    }
}