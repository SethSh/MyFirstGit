using SubmissionCollector.Enums;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models.Segment
{
    internal class RangeMoverValidator
    {
        public ISegmentExcelMatrix ExcelMatrix { get; set; }
        
        public bool Validate()
        {
            var identifier = new SegmentExcelComponentIdentifier();
            if (!identifier.Validate()) return false;

            ExcelMatrix = identifier.ExcelMatrix;
            if (ExcelMatrix is IRangeMovable) return true;

            MessageHelper.Show("Data component not movable", MessageType.Stop);
            return false;
        }

        
    }
}
