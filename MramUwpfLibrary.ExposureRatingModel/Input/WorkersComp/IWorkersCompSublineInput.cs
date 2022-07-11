namespace MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp
{
    public interface IWorkersCompSublineInput : ISublineExposureRatingInput
    {
        WorkersCompCurveSetContainer CurveContainer { get; set; }
    }
}
