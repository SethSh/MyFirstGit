using System;
using System.Collections.Generic;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.Model
{
    public interface IModel
    {
        bool IsDirty { get; set; }
        Guid Guid { get; set; }
        string Name { get; set; }
        long? SourceId { get; set; }
        long? PredecessorSourceId { get; set; }
        DateTime? SourceTimestamp { get; set; }
    }

    
    public interface ISourceComponentModel : IModel
    {
        void BuildQualityControlMessage(StringBuilder stringBuilder, string segmentName);
        void BuildValidationMessage(IList<Allocation> sublineAllocations, StringBuilder stringBuilder, string segmentName);

        int InterDisplayOrder { get; set; }
        int IntraDisplayOrder { get; set; }

    }

    public abstract class BaseSourceModel : IModel
    {
        public bool IsDirty { get; set; }
        public Guid Guid { get; set; }
        public long? SourceId { get; set; }
        public long? PredecessorSourceId { get; set; }
        public DateTime? SourceTimestamp { get; set; }
        public string Name { get; set; }
    }
}
