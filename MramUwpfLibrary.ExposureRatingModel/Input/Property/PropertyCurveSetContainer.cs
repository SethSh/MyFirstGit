using System.Collections.Generic;
using MramUwpfLibrary.ExposureRatingModel.Property;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Property
{
    public class PropertyCurveSetContainer
    {
        public PropertyCurveSetContainer()
        {
            OccupancyTypeAllocations = new List<ProfileAllocation>();
            ConstructionTypeAllocations = new List<ProfileAllocation>();
            ProtectionClassAllocations = new List<ProfileAllocation>();
            TotalInsuredValueAllocations = new List<ITotalInsuredValueAllocation>();
        }
        
        public IEnumerable<IProfileAllocation> OccupancyTypeAllocations { get; set; }
        public IEnumerable<IProfileAllocation> ConstructionTypeAllocations { get; set; }
        public IEnumerable<IProfileAllocation> ProtectionClassAllocations { get; set; }
        public IEnumerable<ITotalInsuredValueAllocation> TotalInsuredValueAllocations { get; set; }
    }
}
