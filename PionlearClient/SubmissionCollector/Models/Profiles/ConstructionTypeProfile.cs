using System.Collections.Generic;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class ConstructionTypeProfile : BaseProfile
    {
        public ConstructionTypeProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new ConstructionTypeExcelMatrix(segmentId);
            Name = BexConstants.ConstructionTypeProfileName;
        }

        [JsonIgnore]
        public ConstructionTypeExcelMatrix ExcelMatrix => (ConstructionTypeExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            return new ConstructionTypeModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                Id = ComponentId,
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
