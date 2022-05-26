using Newtonsoft.Json;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class SublineProfile : BaseProfile
    {
        public SublineProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new SublineExcelMatrix(segmentId);
        }

        [JsonIgnore]
        public SublineExcelMatrix SublineExcelMatrix => (SublineExcelMatrix)CommonExcelMatrix;
        
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

        protected override BaseSourceComponentModel MapToModel()
        {
            throw new System.NotImplementedException();
        }
    }
}
