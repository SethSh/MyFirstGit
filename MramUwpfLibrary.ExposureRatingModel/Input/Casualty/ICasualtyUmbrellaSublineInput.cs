using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{
    public interface ICasualtyUmbrellaSublineInput : ISublineExposureRatingInput
    {
        IEnumerable<UmbrellaTypeAllocation> UmbrellaTypeAllocations { get; set; }
    }
}