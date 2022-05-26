using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class WorkersCompStateHazardGroupProfile : BaseProfile, IProvidesState
    {
        public WorkersCompStateHazardGroupProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new WorkersCompStateHazardGroupExcelMatrix(segmentId);
            Name = BexConstants.WorkersCompStateHazardGroupProfileName;
        }

        [JsonIgnore]
        public WorkersCompStateHazardGroupExcelMatrix ExcelMatrix => (WorkersCompStateHazardGroupExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            var sublines = GetSegment().ToList();

            var model = new WorkersCompStateHazardGroupModel
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
