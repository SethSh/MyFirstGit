using System;
using PionlearClient.Model;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class AggregateLossModelPlus : CollectorApi.AggregateLossModel, IModel
    {
        public int RowId { get; set; }
        public int RowNumber { get; set; }
        public string Name { get; set; }
        public long? SourceId { get; set; }
        public long? PredecessorSourceId { get; set; }
        public bool IsDirty { get; set; }
        public Guid Guid { get; set; }
        public long? DataSetId { get; set; }
        public DateTime? SourceTimestamp { get; set; }
        
        public bool IsEqualTo(CollectorApi.AggregateLossModel otherLoss)
        {
            if (StartDate.DateTime.ToShortDateString() != otherLoss.StartDate.DateTime.ToShortDateString()) return false;
            if (EndDate.DateTime.ToShortDateString() != otherLoss.EndDate.DateTime.ToShortDateString()) return false;

            if (!PaidLossAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PaidLossAmount)) return false;
            if (!PaidAlaeAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PaidAlaeAmount)) return false;
            if (!PaidCombinedAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PaidCombinedAmount)) return false;

            if (!ReportedLossAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.ReportedLossAmount)) return false;
            if (!ReportedAlaeAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.ReportedAlaeAmount)) return false;
            if (!ReportedCombinedAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.ReportedCombinedAmount)) return false;

            return true;
        }
    }
}
