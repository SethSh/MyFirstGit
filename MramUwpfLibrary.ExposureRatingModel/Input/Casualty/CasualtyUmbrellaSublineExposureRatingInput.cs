using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Input.Casualty
{ 
    public class CasualtyUmbrellaSublineExposureRatingInput : SublineExposureRatingInput, ICasualtyUmbrellaSublineInput
    {
        public CasualtyUmbrellaSublineExposureRatingInput(string id):base(id)
        {
            
        }
        
        public IEnumerable<UmbrellaTypeAllocation> UmbrellaTypeAllocations { get; set; }
    }
}