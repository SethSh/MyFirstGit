using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Layers;
using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Input
{
    internal class ReinsuranceParameters
    {
        public ReinsuranceParameters(ILayer layer, ReinsuranceAlaeTreatmentType alaeTreatment, IReinsurancePerspectiveHandler perspectiveHandler)
        {
            Layer = layer;
            AlaeTreatment = alaeTreatment;
            PerspectiveHandler = perspectiveHandler;
        }

        public ReinsuranceParameters(ReinsuranceAlaeTreatmentType alaeTreatment, IReinsurancePerspectiveHandler perspectiveHandler)
        {
            AlaeTreatment = alaeTreatment;
            PerspectiveHandler = perspectiveHandler;
        }

        internal ILayer Layer { get; set; }
        internal ReinsuranceAlaeTreatmentType AlaeTreatment { get; set; }
        internal IReinsurancePerspectiveHandler PerspectiveHandler { get; set; }
    }
}