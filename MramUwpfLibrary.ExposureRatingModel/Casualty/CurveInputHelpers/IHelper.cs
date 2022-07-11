using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers
{
    internal interface IHelper
    {
        CurveInputs GetNumeratorCurveInputs();
        CurveInputs GetNumeratorCurveBenchmarkInputs();
        CurveInputs GetDenominatorCurveInputs();
        CurveInputs GetDenominatorCurveBenchmarkInputs();
    }
}