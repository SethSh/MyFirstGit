using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class WorkersCompStateAttachmentAndWeightPlus : WorkersCompStateAttachmentAndWeight, IProvidesLocation
    {
        public string Location { get; set; }
    }

    internal class WorkersCompStateAttachmentDistributionComparer : IEqualityComparer<WorkersCompStateAttachmentAndWeight>
    {
        public bool Equals(WorkersCompStateAttachmentAndWeight item, WorkersCompStateAttachmentAndWeight otherItem)
        {
            Debug.Assert(item != null, "state attachment != null");
            Debug.Assert(otherItem != null, "other state attachment != null");
            return item.WorkersCompStateId == otherItem.WorkersCompStateId
                   && item.Attachment.IsEpsilonEqual(otherItem.Attachment)
                   && item.Weight.IsEpsilonEqual(otherItem.Weight);
        }

        public int GetHashCode(WorkersCompStateAttachmentAndWeight obj)
        {
            return $"S{obj.WorkersCompStateId}A{obj.Attachment}V{obj.Weight}".GetHashCode();
        }
    }
}