using System.Collections.Generic;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class OccupancyTypeProfile : BaseProfile
    {
        public OccupancyTypeProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new OccupancyTypeExcelMatrix(segmentId);
            Name = BexConstants.OccupancyTypeProfileName;
        }

        
        [JsonIgnore]
        public OccupancyTypeExcelMatrix ExcelMatrix => (OccupancyTypeExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            return new OccupancyTypeModel
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