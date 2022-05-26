using Newtonsoft.Json;
using PionlearClient.Model;
using SubmissionCollector.Models.Historicals.ExcelComponent;

namespace SubmissionCollector.Models.Historicals
{
    public class PeriodSet : BaseHistorical
    {
        public PeriodSet(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new PeriodSetExcelMatrix(SegmentId);
        }

        public override bool IsDirty
        {
            get => base.IsDirty;
            set
            {
                var excelWorkspace = Globals.ThisWorkbook.ThisExcelWorkspace;
                if (excelWorkspace == null) return;

                var segment = excelWorkspace.Package.GetSegment(SegmentId);
                if (segment != null && value) segment.IsDirty = true;

                base.IsDirty = value;
            }
        }
        
        [JsonIgnore]
        public PeriodSetExcelMatrix ExcelMatrix => (PeriodSetExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            throw new System.NotImplementedException();
        }
    }
}
