using System.Collections.Generic;
using Newtonsoft.Json;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class HazardProfile : BaseProfile
    {
        public HazardProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new HazardExcelMatrix(SegmentId);
        }
        
        [JsonIgnore]
        public HazardExcelMatrix ExcelMatrix => (HazardExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            return new HazardModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                SublineCode = ExcelMatrix.Subline.Code,
                Guid = Guid,
                SublineIds = new List<long?> { ExcelMatrix.Subline.Code },
                Items = ExcelMatrix.Items,
                Name = ExcelMatrix.FullName,
                InterDisplayOrder = ExcelMatrix.InterDisplayOrder,
                IntraDisplayOrder = ExcelMatrix.IntraDisplayOrder
            };
        }
    }
}
