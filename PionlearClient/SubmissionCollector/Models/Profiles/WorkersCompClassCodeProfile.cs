using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class WorkersCompClassCodeProfile : BaseProfile
    {
        public WorkersCompClassCodeProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new WorkersCompClassCodeExcelMatrix(segmentId);
            Name = BexConstants.WorkersCompClassCodeProfileName;
        }

        [JsonIgnore]
        public WorkersCompClassCodeExcelMatrix ExcelMatrix => (WorkersCompClassCodeExcelMatrix)CommonExcelMatrix;

        
        protected override BaseSourceComponentModel MapToModel()
        {
            var sublines = GetSegment().ToList();

            var model = new WorkersCompClassCodeModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                Id = ComponentId,
                Guid = Guid,
                SublineIds = sublines.Select(x => new long?(x.Code)).ToList(),
                Items = ExcelMatrix.Items,
                Name = ExcelMatrix.FullName,
            };

            return model;
        }
    }
}
