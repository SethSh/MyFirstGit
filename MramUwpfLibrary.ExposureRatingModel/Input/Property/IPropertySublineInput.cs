namespace MramUwpfLibrary.ExposureRatingModel.Input.Property
{
    public interface IPropertySublineInput : ISublineExposureRatingInput
    {
        PropertyCurveSetContainer CurveContainer { get; set; }
    }
}
