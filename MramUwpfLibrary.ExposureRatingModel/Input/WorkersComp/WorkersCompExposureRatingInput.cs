using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp
{
    public interface IWorkersCompExposureRatingInput : IExposureRatingInput
    {
        IList<IWorkersCompSegmentInput> SegmentInputs { get; set; }
        
    }

    public class WorkersCompExposureRatingInput : BaseExposureRatingInput, IWorkersCompExposureRatingInput
    {
        public WorkersCompExposureRatingInput()
        {
            SegmentInputs = new List<IWorkersCompSegmentInput>();
        }

        public IList<IWorkersCompSegmentInput> SegmentInputs { get; set; }
    }
}