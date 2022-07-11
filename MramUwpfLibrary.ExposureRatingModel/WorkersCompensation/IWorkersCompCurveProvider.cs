using System;
using System.Collections.Generic;
using MramUwpfLibrary.ExposureRatingModel.Casualty;
using MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp;

namespace MramUwpfLibrary.ExposureRatingModel.WorkersCompensation
{
    public interface IWorkersCompCurveProvider
    {
        List<SeverityCurveResult> GetWorkersCompCurve(DateTime effectiveDate, List<WorkersCompStateHazardAllocation> allocation);
        double GetWorkersCompMinnesotaRetentionAmount(DateTime effectiveDate, long minnesotaRetentionCode);
    }

    public class WorkersCompCurveProvider : IWorkersCompCurveProvider
    {
        public List<SeverityCurveResult> GetWorkersCompCurve(DateTime effectiveDate, List<WorkersCompStateHazardAllocation> allocation)
        {
            throw new NotImplementedException();
        }

        public double GetWorkersCompMinnesotaRetentionAmount(DateTime effectiveDate, long minnesotaRetentionCode)
        {
            throw new NotImplementedException();
        }
    }
}
