using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class WorkersCompStateAttachmentProfile : BaseProfile, IProvidesState
    {
        public WorkersCompStateAttachmentProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new WorkersCompStateAttachmentExcelMatrix(segmentId);
            Name = BexConstants.WorkersCompStateAttachmentProfileName;
        }

        [JsonIgnore]
        public WorkersCompStateAttachmentExcelMatrix ExcelMatrix => (WorkersCompStateAttachmentExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            var sublines = GetSegment().ToList();

            var model = new WorkersCompStateAttachmentModel
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
