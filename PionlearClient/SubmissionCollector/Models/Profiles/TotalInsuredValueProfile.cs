using System.Collections.Generic;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class TotalInsuredValueProfile : BaseProfile
    {
        public TotalInsuredValueProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new TotalInsuredValueExcelMatrix(segmentId);
            Name = BexConstants.TotalInsuredValueProfileName;
        }

        [JsonIgnore]
        public TotalInsuredValueExcelMatrix ExcelMatrix => (TotalInsuredValueExcelMatrix)CommonExcelMatrix;
        public bool IsExpanded { get; set; }

        protected override BaseSourceComponentModel MapToModel()
        {
            var model = new TotalInsuredValueModel
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
            
            
            return model;
        }
    }
}
