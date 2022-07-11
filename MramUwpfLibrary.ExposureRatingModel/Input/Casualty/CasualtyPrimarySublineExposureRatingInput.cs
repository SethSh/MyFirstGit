

namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{
    public class CasualtyPrimarySublineExposureRatingInput : SublineExposureRatingInput, ICasualtyPrimarySublineInput
    {
        public CasualtyPrimarySublineExposureRatingInput(string id):base(id)
        {
            CasualtyCurveContainer = new CasualtyCurveSetContainer();
        }

        public CasualtyCurveSetContainer CasualtyCurveContainer { get; set; }
    }
}