using System.Collections.Generic;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class ProtectionClassProfile : BaseProfile
    {
        public ProtectionClassProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new ProtectionClassExcelMatrix(segmentId);
            Name = BexConstants.ProtectionClassProfileName;
        }

        [JsonIgnore]
        public ProtectionClassExcelMatrix ExcelMatrix => (ProtectionClassExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            return new ProtectionClassModel
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