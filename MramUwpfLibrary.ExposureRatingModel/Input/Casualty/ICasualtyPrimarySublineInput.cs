namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{
    public interface ICasualtyPrimarySublineInput : ISublineExposureRatingInput
    {
        CasualtyCurveSetContainer CasualtyCurveContainer { get; set; }
    }
}