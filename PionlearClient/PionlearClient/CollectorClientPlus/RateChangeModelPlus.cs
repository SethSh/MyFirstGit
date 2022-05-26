using System;
using PionlearClient.Model;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class RateChangeModelPlus : CollectorApi.RateChangeModel, IModel
    {
        public int RowId { get; set; }
        public int RowNumber { get; set; }
        public bool IsDirty { get; set; }
        public Guid Guid { get; set; }
        public string Name { get; set; }
        public long? SourceId { get; set; }
        public long? PredecessorSourceId { get; set; }
        public DateTime? SourceTimestamp { get; set; }
        public long? DataSetId { get; set; }

        public bool IsEqualTo(CollectorApi.RateChangeModel otherLoss)
        {
            if (EffectiveDate.Date != otherLoss.EffectiveDate.Date) return false;
            if (!Value.IsEpsilonEqual(otherLoss.Value)) return false;
            
            return true;
        }
    }
}
