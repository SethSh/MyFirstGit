using System.Collections.Generic;
using MramUwpfLibrary.Common.Enums;


namespace MramUwpfLibrary.ExposureRatingModel.Input
{
    public interface ISegmentInput
    {
        string Id { get; set; }
        IList<ISublineExposureRatingInput> SublineInputs { get; set; }
        PolicyAlaeTreatmentType PolicyAlaeTreatment { get; set; }
    }

    public class BaseSegmentInput : ISegmentInput
    {
        public BaseSegmentInput(string id)
        {
            Id = id;
            SublineInputs = new List<ISublineExposureRatingInput>();
        }

        //submission segment id as string
        public string Id { get; set; }
        public IList<ISublineExposureRatingInput> SublineInputs { get; set; }
        public PolicyAlaeTreatmentType PolicyAlaeTreatment { get; set; }
    }
}