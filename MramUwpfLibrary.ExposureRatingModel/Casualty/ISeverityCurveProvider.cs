using System;
using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty
{
    public interface ISeverityCurveProvider
    {
        List<SeverityCurveResult> GetSeverityCurve(DateTime effectiveDate, List<ProfileAllocation> stateProfile,
            long hazardCode, long subLineId);
    }
}
