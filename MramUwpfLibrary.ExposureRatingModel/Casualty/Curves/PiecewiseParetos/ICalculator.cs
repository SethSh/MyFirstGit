using System.Collections.Generic;
using MramUwpfLibrary.Common;
using MramUwpfLibrary.Common.Layers;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.PiecewiseParetos
{
    public interface ICalculator
    {
        double GetLayerLimitedExpectedValue(ILayerPlus layer, Dictionary<int, IParetoParameters> curveParametersSet);
        double GetCdfForLimit(double limit, Dictionary<int, IParetoParameters> curveParametersSetSet);
    }
}
