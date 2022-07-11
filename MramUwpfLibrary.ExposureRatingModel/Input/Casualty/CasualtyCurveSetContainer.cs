using System.Collections.Generic;
using MramUwpfLibrary.Common.Policies;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{
    public class CasualtyCurveSetContainer
    {

        public CasualtyCurveSetContainer()
        {
            StateAllocations = new List<ProfileAllocation>();
            HazardAllocations = new List<ProfileAllocation>();
            PolicySet = new PolicySet();
        }

        public IEnumerable<ProfileAllocation> StateAllocations { get; set; }
        public IEnumerable<ProfileAllocation> HazardAllocations { get; set; }
        public IPolicySet PolicySet { get; set; }
    }
}
