using System.Text;
using PionlearClient.Model;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles
{
    public interface IProfile : IModel
    {

    }

    public abstract class BaseProfile : BaseExcelComponent, IProfile
    {
        protected BaseProfile(int segmentId, int componentId) : base(segmentId, componentId)
        {
            
        }

        protected BaseProfile(int segmentId) : base(segmentId)
        {
        }

        protected ISegment GetSegment()
        {
            var excelWorkspace = Globals.ThisWorkbook.ThisExcelWorkspace;
            return excelWorkspace?.Package.GetSegment(SegmentId);
        }

        public BaseSourceComponentModel CreateModel(StringBuilder validations)
        {
            var validation = CommonExcelMatrix.Validate();
            if (validation.Length > 0)
            {
                validations.AppendLine(CommonExcelMatrix.FullName);
                validations.Append(validation);
                validations.AppendLine();
            }

            var model = MapToModel();
            return model;
        }

        protected abstract BaseSourceComponentModel MapToModel();
    }

}
