namespace MramUwpfLibrary.ExposureRatingModel.Input.Property
{
    public class PropertySublineExposureRatingInput : SublineExposureRatingInput, IPropertySublineInput
    {
        public PropertySublineExposureRatingInput(string id) : base(id)
        {
            CurveContainer = new PropertyCurveSetContainer();
        }

        public PropertyCurveSetContainer CurveContainer { get; set; }
    }
}
