using Newtonsoft.Json;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.Models.DataComponents
{
    public interface IExcelComponent : IModelProxy
    {
        ISegmentExcelMatrix CommonExcelMatrix { get; set; }
        int SegmentId { get; set; }
        int ComponentId { get; set; }
        int InterDisplayOrder { get; }
        int IntraDisplayOrder { get; }
    }

    public abstract class BaseExcelComponent : BaseModelProxy, IExcelComponent
    {
        protected BaseExcelComponent(int segmentId, int componentId)
        {
            SegmentId = segmentId;
            ComponentId = componentId;
        }

        protected BaseExcelComponent(int segmentId)
        {
            SegmentId = segmentId;
        }

        public ISegmentExcelMatrix CommonExcelMatrix { get; set; }
        public int SegmentId { get; set; }
        public int ComponentId { get; set; }

        [JsonIgnore]
        public int IntraDisplayOrder => CommonExcelMatrix.IntraDisplayOrder;

        [JsonIgnore]
        public int InterDisplayOrder => CommonExcelMatrix.InterDisplayOrder;
    }
    
}

