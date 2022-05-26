using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Segment.DataComponents;


namespace SubmissionCollector.ExcelUtilities.RangeTransposer
{
    public abstract class BaseRangeTransposer : IRangeTransposer
    {
        public Range BodyHeaderRange { get; set; }
        public Range HeaderRange { get; set; }
        public Range BodyRange { get; set; }
        public IExcelMatrix ExcelMatrix { get; set; }
        public ISegment Segment { get; set; }
        
        protected BaseRangeTransposer(ISegment segment, IExcelMatrix excelMatrix)
        {
            Segment = segment;
            ExcelMatrix = excelMatrix;

            BodyRange = ((ISegmentExcelMatrix)ExcelMatrix).GetBodyRange();
            BodyHeaderRange = ((ISegmentExcelMatrix)ExcelMatrix).GetBodyHeaderRange();
            HeaderRange = ((ISegmentExcelMatrix)ExcelMatrix).GetHeaderRange();
        }

        public void TransposeWrapper()
        {
            using (new StatusBarUpdater("Transposing ..."))
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    using (new CalculationSetter())
                    {
                        Transpose();
                    }
                }
            }
        }

        public abstract void Transpose();
        
    }
}