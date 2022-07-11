namespace MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp
{
    public class WorkersCompSublineExposureRatingInput : SublineExposureRatingInput, IWorkersCompSublineInput
    {
        public WorkersCompSublineExposureRatingInput(string id) : base(id)
        {
            CurveContainer = new WorkersCompCurveSetContainer();
        }

        public WorkersCompCurveSetContainer CurveContainer { get; set; }
    }
}