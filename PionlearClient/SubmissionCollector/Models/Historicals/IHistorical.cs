using System.Text;
using PionlearClient.Model;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Historicals
{
    public interface IHistorical : IModelProxy
    {
        
    }

    public abstract class BaseHistorical : BaseExcelComponent, IHistorical
    {
        protected BaseHistorical(int segmentId, int componentId) : base(segmentId, componentId)
        {
        }

        protected BaseHistorical(int segmentId) : base(segmentId)
        {
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
