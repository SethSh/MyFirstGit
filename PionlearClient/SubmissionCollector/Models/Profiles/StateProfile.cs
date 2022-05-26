using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class StateProfile : BaseProfile, IProvidesState
    {
        public StateProfile(int segmentId, int componentId) : base(segmentId, componentId)
        {
            CommonExcelMatrix = new StateExcelMatrix(segmentId, componentId);
            Name = BexConstants.StateProfileName;
        }

        [JsonIgnore]
        public StateExcelMatrix ExcelMatrix => (StateExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            return new StateModel
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
        }
    }

    public interface IProvidesState
    {
    }
}
