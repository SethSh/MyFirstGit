using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using MramUwpfLibrary.Common;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Layers;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.PiecewiseParetos
{
    public static class PiecewiseParetoHelper
    {
        public static BreakpointResult FitPiecewisePareto(ILayerLossEstimate layerLossEstimate, double frequencyBottom, double frequencyTop)
        {
            var layerTop = layerLossEstimate.Limit + layerLossEstimate.Attachment;
            if (layerLossEstimate.Attachment >= layerTop || frequencyTop > frequencyBottom)
            {
                if (frequencyTop <= frequencyBottom * 1.000000001)
                {
                    frequencyTop = frequencyBottom;
                }
                else
                {
                    throw new ArgumentException($"Fitting {StringConstants.PiecewiseParetoName.ToLower()} inputs don't make sense");
                }
            }

            var rateOnLine = layerLossEstimate.Loss / layerLossEstimate.Limit;
            double breakpoint;
            if (frequencyBottom < rateOnLine * 1.0000001)
            {
                breakpoint = layerTop - 0.01 * layerLossEstimate.Limit;
                var alpha1 = 0d;
                var alpha2 = Math.Log(frequencyTop / frequencyBottom) / Math.Log(breakpoint / layerTop);
                return new BreakpointResult { Breakpoint = breakpoint, Alpha1 = alpha1, Alpha2 = alpha2 };
            }

            if (frequencyTop > rateOnLine * 0.9999999)
            {
                breakpoint = layerLossEstimate.Attachment + 0.01 * layerLossEstimate.Limit;
                var alpha1 = Math.Log(frequencyTop / frequencyBottom) / Math.Log(layerLossEstimate.Attachment / breakpoint);
                var alpha2 = 0d;
                return new BreakpointResult { Breakpoint = breakpoint, Alpha1 = alpha1, Alpha2 = alpha2 };
            }

            breakpoint = CalculateBreakpoint(layerLossEstimate, frequencyBottom, frequencyTop);
            var alphaResult = CalculateAlphas(layerLossEstimate, breakpoint, frequencyBottom, frequencyTop);
            if (alphaResult.Code == 0)
            {
                return new BreakpointResult { Breakpoint = breakpoint, Alpha1 = alphaResult.Alpha1, Alpha2 = alphaResult.Alpha2 };
            }

            throw new ArgumentException($"Fitting {StringConstants.PiecewiseParetoName.ToLower()} failed");
        }

        public static double GetParetoPieces(double alpha, ILayerLossEstimate layer, double breakPoint, double frequencyBottom, double frequencyTop)
        {
            var alpha2 = GetAlphaTwo(alpha, layer, breakPoint, frequencyBottom, frequencyTop);

            return frequencyBottom * GetExpectedLayerLoss(layer.Bottom, breakPoint, alpha)
                   + frequencyBottom * Math.Pow(layer.Bottom / breakPoint, alpha) 
                   * GetExpectedLayerLoss(breakPoint, layer.Top, alpha2);
        }

        public static double GetExpectedLayerLoss(ILayer layer, double alpha)
        {
            var oneMinusAlpha = 1 - alpha;

            if (layer.Attachment.IsEqualToZero() || layer.Limit < 0 || alpha < 0)
            {
                throw new ArgumentException("Can't calculate piecewise pareto layer expected value ");
            }

            var layerTop = layer.Limit+layer.Attachment;
            if (!oneMinusAlpha.IsEqualToZero())
            {
                return layer.Attachment / oneMinusAlpha * (Math.Pow(layerTop / layer.Attachment, oneMinusAlpha) - 1);
            }

            if (alpha <= 10000)
            {
                return layer.Attachment * (Math.Log(layerTop) - Math.Log(layer.Attachment));
            }

            return 0;
        }

        public static double GetExpectedLayerLoss(double layerBottom, double layerTop, double alpha)
        {
            return GetExpectedLayerLoss(new ExcessLayer(limit: layerTop - layerBottom, attachment: layerBottom), alpha);
        }

        public static AlphaResult CalculateAlphas(ILayerLossEstimate layerLossEstimate, double breakPoint, double frequencyBottom, double frequencyTop)
        {
            var rateOnLine = layerLossEstimate.Loss / layerLossEstimate.Limit;

            if (rateOnLine > frequencyBottom || rateOnLine < frequencyTop) return new AlphaResult { Code = 3 };

            if (rateOnLine.IsEqualTo(frequencyBottom) || rateOnLine.IsEqualTo(frequencyTop) && frequencyBottom.IsNotEqualTo(frequencyTop))
            {
                return new AlphaResult { Code = 3 };
            }

            if (rateOnLine.IsEqualTo(frequencyBottom) && rateOnLine.IsEqualTo(frequencyTop)) return new AlphaResult { Alpha1 = 0, Alpha2 = 0, Code = 0 };

            var layerTop = layerLossEstimate.Top;
            if (layerTop <= layerLossEstimate.Attachment || layerLossEstimate.Attachment <= 0 || breakPoint <= layerLossEstimate.Attachment ||
                breakPoint >= layerTop)
            {
                return new AlphaResult { Code = 4 };
            }

            var maximumExpectedLoss = GetParetoPieces(0, layerLossEstimate, breakPoint, frequencyBottom, frequencyTop);
            var alpha2 = 0d;
            var alpha1 = GetAlphaOne(0, layerLossEstimate, breakPoint, frequencyBottom, frequencyTop);
            var minimumExpectedLoss = GetParetoPieces(alpha1, layerLossEstimate, breakPoint, frequencyBottom, frequencyTop);

            if (layerLossEstimate.Loss > maximumExpectedLoss) return new AlphaResult { Alpha1 = alpha1, Alpha2 = alpha2, Code = 2 };
            if (layerLossEstimate.Loss < minimumExpectedLoss) return new AlphaResult { Alpha1 = alpha1, Alpha2 = alpha2, Code = 1 };

            //bisection
            var upperBoundAlpha1 = alpha1;
            var lowerBoundAlpha1 = 0d;

            const double precision = 1e-7;
            while (upperBoundAlpha1 - lowerBoundAlpha1 >= precision)
            {
                alpha1 = (upperBoundAlpha1 + lowerBoundAlpha1) / 2;
                var temp = GetParetoPieces(alpha1, layerLossEstimate, breakPoint, frequencyBottom, frequencyTop);

                if (temp > layerLossEstimate.Loss)
                {
                    lowerBoundAlpha1 = alpha1;
                }
                else if (temp <= layerLossEstimate.Loss)
                {
                    upperBoundAlpha1 = alpha1;
                }
                else
                {
                    return new AlphaResult { Alpha1 = alpha1, Alpha2 = alpha2, Code = 4 };
                }
            } 


            alpha1 = (upperBoundAlpha1 + lowerBoundAlpha1) / 2;
            alpha2 = GetAlphaTwo(alpha1, layerLossEstimate, breakPoint, frequencyBottom, frequencyTop);
            return new AlphaResult { Alpha1 = alpha1, Alpha2 = alpha2, Code = 0 };
        }

        public static Dictionary<int, double> GetSurvivalValues(Dictionary<int, IParetoParameters> sortedCurveParametersSet)
        {
            var parametersSetIds = sortedCurveParametersSet.Keys;
            var survivals = new Dictionary<int, double>();
            var firstId = parametersSetIds.First();
            foreach (var id in parametersSetIds)
            {
                survivals.Add(id, 1d);
            }

            var previousKey = 0;
            foreach (var parameters in sortedCurveParametersSet)
            {
                if (parameters.Key == firstId)
                {
                    continue;
                }

                var previousSurvival = survivals[previousKey];
                var previousParameters = sortedCurveParametersSet[previousKey];
                var amount = previousSurvival *
                             Math.Pow(previousParameters.Threshold / parameters.Value.Threshold, previousParameters.Alpha);

                survivals[parameters.Key] = amount;
                previousKey = parameters.Key;
            }
            
            return survivals;
        }

        public static double FindAlpha(ILayerLossEstimate layerLossEstimate1, ILayerLossEstimate layerLossEstimate2)
        {
            var layer1Top = layerLossEstimate1.Limit + layerLossEstimate1.Attachment;
            var layer2Top = layerLossEstimate2.Limit + layerLossEstimate2.Attachment;

            if (layerLossEstimate1.Limit <= 0 || layerLossEstimate1.Attachment <= 0 || layerLossEstimate2.Limit <= 0 ||
                layerLossEstimate2.Attachment <= 0)
                throw new ArgumentOutOfRangeException();

            if (layer1Top <= layer2Top && layerLossEstimate1.Limit <= layerLossEstimate2.Limit &&
                layerLossEstimate1.Attachment >= layerLossEstimate2.Attachment && layerLossEstimate1.Loss < layerLossEstimate2.Loss)
                throw new ArgumentOutOfRangeException();

            if (layer1Top >= layer2Top && layerLossEstimate1.Limit >= layerLossEstimate2.Limit &&
                layerLossEstimate1.Attachment <= layerLossEstimate2.Attachment && layerLossEstimate1.Loss > layerLossEstimate2.Loss)
                throw new ArgumentOutOfRangeException();


            var relLoss = layerLossEstimate2.Loss / layerLossEstimate1.Loss;
            var relCovers = layerLossEstimate2.Limit / layerLossEstimate1.Limit;
            var relLog = Math.Log(1 + (layerLossEstimate2.Limit / layerLossEstimate2.Attachment)) /
                         Math.Log(1 + (layerLossEstimate1.Limit / layerLossEstimate1.Attachment));

            const double eps = 0.0001;
            if (relLoss.IsEpsilonEqual(relLog, eps))
                return 1;

            if (layerLossEstimate2.Loss.IsEpsilonEqual(layerLossEstimate1.Loss * relCovers, eps))
                return 0;

            double alphaLeft;
            double alphaRight;
            double left;

            if ((relCovers < relLoss && relLoss < relLog) || (relCovers > relLoss && relLoss > relLog))
            {
                alphaLeft = 0;
                alphaRight = 1;
                left = relCovers;
            }
            else if ((relCovers < relLog && relLog < relLoss) || (relCovers > relLog && relLog > relLoss))
            {
                alphaLeft = 1;
                alphaRight = 30;
                left = relLog;
                ExtrapolatePareto(layerLossEstimate1, layerLossEstimate2, alphaRight);
            }
            else
            {
                alphaLeft = -10;
                alphaRight = 0;
                left = ExtrapolatePareto(layerLossEstimate1, layerLossEstimate2, alphaLeft);
            }

            double result = 0;
            while (alphaRight - alphaLeft >= eps)
            {
                result = 0.5 * (alphaLeft + alphaRight);
                var midValue = ExtrapolatePareto(layerLossEstimate1, layerLossEstimate2, result);
                if ((left < relLoss && midValue < relLoss) || (left > relLoss && midValue > relLoss))
                {
                    alphaLeft = result;
                    left = midValue;
                }
                else
                {
                    alphaRight = result;
                }
            }

            if (result < 0)
            {
                throw new OverflowException();
            }
            return result;
        }

        private static double ExtrapolatePareto(ILayerLossEstimate layer1, ILayerLossEstimate layer2, double alpha)
        {
            double result;
            if (layer1.Limit <= 0 || layer1.Attachment <= 0 || layer2.Limit <= 0 ||
                layer2.Attachment <= 0)
                throw new ArgumentOutOfRangeException();
            if (alpha.IsEqualTo(1))
            {
                result = Math.Log(layer2.Top / layer2.Bottom) /
                         Math.Log(layer1.Top / layer1.Bottom);
            }
            else
            {
                result = (Math.Pow(layer2.Top, (1 - alpha)) - Math.Pow(layer2.Bottom, (1 - alpha))) /
                         (Math.Pow(layer1.Top, (1 - alpha)) - Math.Pow(layer1.Bottom, (1 - alpha)));
            }
            return result;
        }

        private static double GetAlphaOne(double iAlpha, ILayerLossEstimate layer, double breakPoint, double freqBottom, double freqTop)
        {
            var alpha = (Math.Log(freqTop / freqBottom) - iAlpha * Math.Log(breakPoint/layer.Top)) 
                / Math.Log(layer.Bottom/breakPoint);
            if (alpha >= 0) return alpha;

            if (alpha > -1e-7)
            {
                alpha = 0;
            }
            else
            {
                throw new Exception("Error calculating alpha.");
            }
            return alpha;
        }

        private static double GetAlphaTwo(double iAlpha, ILayerLossEstimate layer, double breakPoint, double freqBottom, double freqTop)
        {
            var alpha = (Math.Log(freqTop / freqBottom) - iAlpha * Math.Log(layer.Bottom/breakPoint)) 
                / Math.Log(breakPoint/layer.Top);
            if (alpha >= 0) return alpha;

            if (alpha > -1e-7)
            {
                alpha = 0;
            }
            else
            {
                throw new Exception("Error calculating alpha.");
            }
            return alpha;
        }

        private static double CalculateBreakpoint(ILayerLossEstimate layerLossEstimate, double frequencyBottom, double frequencyTop)
        {
            var breakpoint = (layerLossEstimate.Bottom + layerLossEstimate.Top) / 2; //Average
            var maximumRatio = 0d;
            double ratio;
            var delta = layerLossEstimate.Limit / 4;

            var alphaResult = CalculateAlphas(layerLossEstimate, breakpoint, frequencyBottom, frequencyTop);
            if (alphaResult.Code == 0)
            {
                ratio = GetAlphaRatio(alphaResult.Alpha1, alphaResult.Alpha2);
                if (ratio > maximumRatio) maximumRatio = ratio;
            }

            var i = 0;
            while (i <= 100 && (i <= 10 || alphaResult.Code != 0))
            {
                if (alphaResult.Code == 1)
                {
                    breakpoint -= delta;
                    alphaResult = CalculateAlphas(layerLossEstimate, breakpoint, frequencyBottom, frequencyTop);
                    if (alphaResult.Code == 0)
                    {
                        ratio = GetAlphaRatio(alphaResult.Alpha1, alphaResult.Alpha2);
                        if (ratio > maximumRatio) maximumRatio = ratio;
                    }
                }
                else if (alphaResult.Code == 2)
                {
                    breakpoint += delta;
                    alphaResult = CalculateAlphas(layerLossEstimate, breakpoint, frequencyBottom, frequencyTop);
                    if (alphaResult.Code == 0)
                    {
                        ratio = GetAlphaRatio(alphaResult.Alpha1, alphaResult.Alpha2);
                        if (ratio > maximumRatio) maximumRatio = ratio;
                    }
                }
                else if (alphaResult.Code == 0)
                {
                    var isFinished = false;
                    var temp = breakpoint + delta;
                    alphaResult = CalculateAlphas(layerLossEstimate, temp, frequencyBottom, frequencyTop);
                    if (alphaResult.Code == 0)
                    {
                        ratio = GetAlphaRatio(alphaResult.Alpha1, alphaResult.Alpha2);
                        if (ratio > maximumRatio)
                        {
                            maximumRatio = ratio;
                            breakpoint = temp;
                            isFinished = true;
                        }
                    }
                    if (!isFinished)
                    {
                        temp = breakpoint - delta;
                        alphaResult = CalculateAlphas(layerLossEstimate, temp, frequencyBottom, frequencyTop);
                        if (alphaResult.Code == 0)
                        {
                            ratio = GetAlphaRatio(alphaResult.Alpha1, alphaResult.Alpha2);
                            if (ratio > maximumRatio)
                            {
                                maximumRatio = ratio;
                                breakpoint = temp;
                                isFinished = true; 
                            }
                        }
                    }
                }

                delta /= 2;
                i++;

            }

            if (alphaResult.Code == 0) return breakpoint;
            throw new InvalidConstraintException($"Couldn't find {StringConstants.PiecewiseParetoName.ToLower()} breakpoint");
        }

        private static double GetAlphaRatio(double alpha1, double alpha2)
        {
            if (alpha1 > alpha2) return alpha2 / alpha1;

            if (alpha2.IsEqualToZero()) return 0;

            return alpha1 / alpha2;
        }

    }
}
