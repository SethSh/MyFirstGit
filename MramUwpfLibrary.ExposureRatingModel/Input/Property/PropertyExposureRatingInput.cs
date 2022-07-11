using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Property
{
    public interface IPropertyExposureRatingInput : IExposureRatingInput
    {
        IList<IPropertySegmentInput> SegmentInputs { get; set; }
    }
    public class PropertyExposureRatingInput : BaseExposureRatingInput, IPropertyExposureRatingInput
    {
        public PropertyExposureRatingInput()
        {
            SegmentInputs = new List<IPropertySegmentInput>();
        }
        public IList<IPropertySegmentInput> SegmentInputs { get; set; }
    }
}
