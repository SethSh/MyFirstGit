using System;
using MramUwpfLibrary.Common.Enums;

namespace MramUwpfLibrary.ExposureRatingModel
{
    internal static class Extensions
    {
        internal static ReinsuranceAlaeTreatmentType MapToReinsuranceAlaeTreatmentType(
            this LossRatioAlaeTreatmentType lossRatioAlaeTreatmentType)
        {
            switch (lossRatioAlaeTreatmentType)
            {
                case LossRatioAlaeTreatmentType.PartOfLoss:
                    return ReinsuranceAlaeTreatmentType.PartOfLoss;

                case LossRatioAlaeTreatmentType.ProRata: 
                    return ReinsuranceAlaeTreatmentType.ProRata;
                
                case LossRatioAlaeTreatmentType.Excluded:
                    return ReinsuranceAlaeTreatmentType.Excluded;
                    
                default:
                    throw new ArgumentOutOfRangeException(nameof(lossRatioAlaeTreatmentType), lossRatioAlaeTreatmentType, null);
            }
        }

        internal static ReinsuranceAlaeTreatmentType MapToReinsuranceAlaeTreatmentTypeForLossRatio(
            this PolicyAlaeTreatmentType policyAlaeTreatmentType)
        {
            switch (policyAlaeTreatmentType)
            {
                case PolicyAlaeTreatmentType.InAdditionToLimit: 
                    return ReinsuranceAlaeTreatmentType.ProRata;

                case PolicyAlaeTreatmentType.WithinLimit: 
                    return ReinsuranceAlaeTreatmentType.PartOfLoss;
                    
                default:
                    throw new ArgumentOutOfRangeException(nameof(policyAlaeTreatmentType), policyAlaeTreatmentType, null);
            }
        }

        internal static ReinsuranceAlaeTreatmentType MapToReinsuranceAlaeTreatmentTypeForLayer(
            this PolicyAlaeTreatmentType policyAlaeTreatmentType, ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment)
        {
            switch (policyAlaeTreatmentType)
            {
                case PolicyAlaeTreatmentType.WithinLimit:
                    return ReinsuranceAlaeTreatmentType.PartOfLoss;
                
                case PolicyAlaeTreatmentType.InAdditionToLimit:
                    return reinsuranceAlaeTreatment == ReinsuranceAlaeTreatmentType.Excluded 
                        ? ReinsuranceAlaeTreatmentType.Excluded 
                        : ReinsuranceAlaeTreatmentType.ProRata;

                default:
                    throw new ArgumentOutOfRangeException(nameof(policyAlaeTreatmentType), policyAlaeTreatmentType, null);
            }
        }

        internal static ReinsuranceAlaeTreatmentType MapToReinsuranceAlaeTreatmentTypeForUnlimitedLossRatio(this PolicyAlaeTreatmentType policyAlaeTreatmentType)
        {
            switch (policyAlaeTreatmentType)
            {
                case PolicyAlaeTreatmentType.WithinLimit:
                    return ReinsuranceAlaeTreatmentType.PartOfLoss;
                
                case PolicyAlaeTreatmentType.InAdditionToLimit:
                    return ReinsuranceAlaeTreatmentType.ProRata;

                default:
                    throw new ArgumentOutOfRangeException(nameof(policyAlaeTreatmentType), policyAlaeTreatmentType, null);
            }
        }
    }
}

