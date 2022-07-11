using System.Collections.Generic;
using System.Linq;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Property
{
    public interface IPropertySegmentInput : ISegmentInput
    {
        
        IEnumerable<PropertySublineExposureRatingInput> PrimaryInputs { get; }
    }
    public class PropertySegmentInput : BaseSegmentInput, IPropertySegmentInput
    {
        public PropertySegmentInput(string id) : base(id)
        {

        }
        public IEnumerable<PropertySublineExposureRatingInput> PrimaryInputs => SublineInputs.OfType<PropertySublineExposureRatingInput>();
    }

}
