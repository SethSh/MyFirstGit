using System.Collections.Generic;
using System.Linq;

namespace MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp
{
    public interface IWorkersCompSegmentInput : ISegmentInput
    {
        IEnumerable<WorkersCompSublineExposureRatingInput> PrimaryInputs { get; }
    }

    public class WorkersCompSegmentInput : BaseSegmentInput, IWorkersCompSegmentInput
    {
        public WorkersCompSegmentInput(string id) : base(id)
        {
        
        }

        public IEnumerable<WorkersCompSublineExposureRatingInput> PrimaryInputs => SublineInputs.OfType<WorkersCompSublineExposureRatingInput>();
    }
}