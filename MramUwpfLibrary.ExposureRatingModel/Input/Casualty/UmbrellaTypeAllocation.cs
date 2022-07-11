namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{
    public class UmbrellaTypeAllocation
    {
        public UmbrellaTypeAllocation()
        {
            CasualtyCurveSetContainer = new CasualtyCurveSetContainer();
        }
        
        public double Allocation { get; set; }
        public double NormalizedAllocation { get; set; }
        public CasualtyCurveSetContainer CasualtyCurveSetContainer { get; set; }
        public bool IsPersonal { get; set; }
    }
}