using System;
using PionlearClient.Model;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class ExposureModelPlus : CollectorApi.ExposureModel, IModel
    {
        public int RowId { get; set; }
        public string Location { get; set; }
        public string Name { get; set; }
        public long? SourceId { get; set; }
        public long? PredecessorSourceId { get; set; }
        public bool IsDirty { get; set; }
        public Guid Guid { get; set; }
        public long? DataSetId { get; set; }
        public DateTime? SourceTimestamp { get; set; }
        
        public bool IsEqualTo(CollectorApi.ExposureModel otherExposure)
        {
            if (StartDate.DateTime.ToShortDateString() != otherExposure.StartDate.DateTime.ToShortDateString()) return false;
            if (EndDate.DateTime.ToShortDateString() != otherExposure.EndDate.DateTime.ToShortDateString()) return false;

            if (!Amount.IsEqual(otherExposure.Amount)) return false;
            return true;
        }
    }
}
