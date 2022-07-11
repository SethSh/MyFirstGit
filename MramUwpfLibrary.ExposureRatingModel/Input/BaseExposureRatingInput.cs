using System;
using System.Collections.Generic;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Layers;
using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Input
{
    public interface IExposureRatingInput
    {
        IEnumerable<ILayer> ReinsuranceLayers { get; set; }
        ReinsuranceAlaeTreatmentType ReinsuranceAlaeTreatment { get; set; }
        IReinsurancePerspectiveHandler ReinsurancePerspective { get; set; }
        bool UseAlternative { get; set; }
        DateTime EffectiveDate { get; set; }
    }

    public class BaseExposureRatingInput : IExposureRatingInput
    {
        private bool _isLayerWithRespectToAttachment; 
        public ReinsuranceAlaeTreatmentType ReinsuranceAlaeTreatment { get; set; }
        public IReinsurancePerspectiveHandler ReinsurancePerspective { get; set; }
        public IEnumerable<ILayer> ReinsuranceLayers { get; set; }
        public bool UseAlternative { get; set; }
        public DateTime EffectiveDate { get; set; }

        public bool IsLayerWithRespectToAttachment
        {
            get => _isLayerWithRespectToAttachment;
            set
            {
                _isLayerWithRespectToAttachment = value;
                ReinsurancePerspective = ReinsurancePerspectiveFactory.Create(_isLayerWithRespectToAttachment);
            }
        }
    }
}