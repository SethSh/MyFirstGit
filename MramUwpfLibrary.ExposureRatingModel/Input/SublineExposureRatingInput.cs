using MramUwpfLibrary.Common.Enums;

namespace MramUwpfLibrary.ExposureRatingModel.Input
{
    public interface ISublineExposureRatingInput
    {
        string Id { get; }
        string SublineCode { get; }
        string GroupId { get; set; }
        double AlaeAdjustmentFactor { get; set; }
        double LimitedLossRatio { get; set; }
        LossRatioAlaeTreatmentType LossRatioAlaeTreatment { get; set; }
        double LossRatioLimit { get; set; }
        double GroupExposureAmount { get; set; }
        double Allocation { get; set; }
        double AllocatedExposureAmount { get; }
    }

    public abstract class SublineExposureRatingInput : ISublineExposureRatingInput
    {
        protected SublineExposureRatingInput(string id)
        {
            Id = id;
        }

        public string Id { get; } //id is subline id
        public string SublineCode { get; set; }

        public string GroupId { get; set; } //group is submission segment
        public double GroupExposureAmount { get; set; }

        public double Allocation { get; set; }
        public double NormalizedAllocation { get; set; }
        public double AllocatedExposureAmount => GroupExposureAmount * NormalizedAllocation;

        public double AlaeAdjustmentFactor { get; set; }
        public double LimitedLossRatio { get; set; }
        public double LossRatioLimit { get; set; }
        public LossRatioAlaeTreatmentType LossRatioAlaeTreatment { get; set; }

        public bool IsPersonal { get; set; }
    }
}