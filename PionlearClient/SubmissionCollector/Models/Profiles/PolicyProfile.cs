using System;
using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class PolicyProfile : BaseProfile
    {
        public PolicyProfile(int segmentId, int componentId) : base(segmentId, componentId)
        {
            CommonExcelMatrix = new PolicyExcelMatrix(segmentId, componentId);
            Name = BexConstants.PolicyProfileName;
        }

        public int? UmbrellaType { get; set; }

        [JsonIgnore]
        public PolicyExcelMatrix ExcelMatrix => (PolicyExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            var model = new PolicyModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                Id = ComponentId,
                Guid = Guid,
                SublineIds = ExcelMatrix.Select(x => new long?(x.Code)).ToList(),
                Items = ExcelMatrix.Items,
                Name = ExcelMatrix.FullName,
                InterDisplayOrder = ExcelMatrix.InterDisplayOrder,
                IntraDisplayOrder = ExcelMatrix.IntraDisplayOrder
            };

            var segment = GetSegment();

            if (segment.IsUmbrella) // Personal is not an umbrella type - patch here to get correct umbrella type code
            {
                model.UmbrellaTypeId = segment.ContainsAnyCommercialSublines
                    ? UmbrellaType
                    : Convert.ToInt16(UmbrellaTypesFromBex.GetPersonalCode());
            }
            return model;
        }
    }
}
