using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp
{
    public class WorkersCompCurveSetContainer
    {
        public WorkersCompCurveSetContainer()
        {
            StateHazardGroupAllocations = new List<WorkersCompStateHazardAllocation>();
            StateAttachmentAllocations = new List<WorkersCompStateAttachmentAllocation>();
        }

        //mandatory
        public IEnumerable<WorkersCompStateHazardAllocation> StateHazardGroupAllocations { get; set; }
        
        //optional
        public IEnumerable<WorkersCompStateAttachmentAllocation> StateAttachmentAllocations { get; set; }

        public double MinnesotaRetentionValue { get; set; }

        //get wc severity curve - based on [date  state id  hazard id]
    }

    public class WorkersCompStateHazardAllocation
    {
        public long StateId { get; set; }
        public long HazardId { get; set; }
        public double Value { get; set; }
    }

    public class WorkersCompStateAttachmentAllocation
    {
        public long StateId { get; set; }
        public double Attachment { get; set; }
        public double Value { get; set; }
    }

}