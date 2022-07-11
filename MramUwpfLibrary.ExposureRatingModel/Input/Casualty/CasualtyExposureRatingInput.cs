using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{
    //used for submission segments

    public interface ICasualtyExposureRatingInput : IExposureRatingInput
    {
        IList<ICasualtySegmentInput> SegmentInputs { get; set; }
    }

    public class CasualtyExposureRatingInput : BaseExposureRatingInput, ICasualtyExposureRatingInput
    {
        public CasualtyExposureRatingInput()
        {
            SegmentInputs = new List<ICasualtySegmentInput>();
        }
        public IList<ICasualtySegmentInput> SegmentInputs { get; set; }
    }
}