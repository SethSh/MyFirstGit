using System.Collections.Generic;
using System.Linq;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{

    public interface ICasualtySegmentInput : ISegmentInput
    {
     
        IEnumerable<CasualtyPrimarySublineExposureRatingInput> PrimaryInputs { get; }
      
        IEnumerable<CasualtyUmbrellaSublineExposureRatingInput> UmbrellaInputs { get; }
    }

    public class CasualtySegmentInput : BaseSegmentInput, ICasualtySegmentInput
    {
        public CasualtySegmentInput(string id) : base(id)
        {

        }
        public IEnumerable<CasualtyPrimarySublineExposureRatingInput> PrimaryInputs => SublineInputs.OfType<CasualtyPrimarySublineExposureRatingInput>();
        public IEnumerable<CasualtyUmbrellaSublineExposureRatingInput> UmbrellaInputs => SublineInputs.OfType<CasualtyUmbrellaSublineExposureRatingInput>();
    }
}
