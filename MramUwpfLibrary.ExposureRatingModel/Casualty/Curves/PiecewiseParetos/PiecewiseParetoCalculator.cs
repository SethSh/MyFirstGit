using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using MramUwpfLibrary.Common;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Layers;
using MramUwpfLibrary.ExposureRatingModel.Discretize;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.PiecewiseParetos
{
    public class PiecewiseParetoCalculatorSequel : ICalculator
    {
        public double GetLayerLimitedExpectedValue(ILayerPlus layer, Dictionary<int, IParetoParameters> curveParametersSet)
        {
            throw new NotImplementedException();
        }

        public double GetCdfForLimit(double limit, Dictionary<int, IParetoParameters> curveParametersSetSet)
        {
            throw new NotImplementedException();
        }

        public IDiscretization Discretize(IDiscretizeInput discretizeInput, Dictionary<int, IParetoParameters> curveParametersSet)
        {
            var cumulativeDiscretization = BuildShell(discretizeInput);

            var cps = curveParametersSet.OrderBy(cp => cp.Value.Threshold);
            var counter = 0;
            var sortedCurveSet = cps.ToDictionary(pair => counter++, pair => pair.Value);
            var maximumCurveIndex = sortedCurveSet.Last().Key;

            foreach (var dictionaryItem in sortedCurveSet)
            {
                var curve = dictionaryItem.Value;
                List<IDiscretizationItem> distributionItems;
                if (dictionaryItem.Key < maximumCurveIndex)
                {
                    var nextCurve = sortedCurveSet[dictionaryItem.Key + 1];
                    distributionItems = cumulativeDiscretization.Where(d => (d.Loss + discretizeInput.BucketWidth * 0.5) >= dictionaryItem.Value.Threshold 
                    && (d.Loss + discretizeInput.BucketWidth * 0.5) < nextCurve.Threshold).ToList();
                }
                else
                {
                    distributionItems = cumulativeDiscretization.Where(d => (d.Loss + discretizeInput.BucketWidth * 0.5) >= dictionaryItem.Value.Threshold).ToList();
                }

                if (!distributionItems.Any()) continue;

                var curveSetLessThanOrEqualToThisOne = sortedCurveSet.Where(cs => cs.Key <= dictionaryItem.Key).ToList();
                var cumulativeSurvival = GetCumulativeSurvival(curveSetLessThanOrEqualToThisOne);

                foreach (var item in distributionItems)
                {
                    var loss = item.Loss + discretizeInput.BucketWidth * 0.5;
                    var survival = Math.Pow(curve.Threshold / loss, curve.Alpha);
                    item.Probability = 1 - survival * cumulativeSurvival;
                }
            }

            cumulativeDiscretization.Last().Probability = 1;

            var discretization = cumulativeDiscretization.MapCumulativeToIncremental();
            return discretization;
        }

        private static double GetCumulativeSurvival(IList<KeyValuePair<int, IParetoParameters>> curveSet)
        {
            var amount = 1d;
            foreach (var pair in curveSet.Where(cs => cs.Key > 0))
            {
                var previous = curveSet[pair.Key - 1];
                amount *= Math.Pow(previous.Value.Threshold / pair.Value.Threshold, previous.Value.Alpha);
            }
            return amount;
        }

        private static Discretization BuildShell(IDiscretizeInput discretizeInput)
        {
            var count = discretizeInput.BucketCount;
            var width = discretizeInput.BucketWidth;

            var discretization = new Discretization();
            var ns = Enumerable.Range(0, count);
            
            foreach (var n in ns)
            {
                discretization.Add(new DiscretizationItem { Loss = n * width, Probability = 0 });
            }

            return discretization;
        }
    }

    public class PiecewiseParetoCalculator : ICalculator
    {
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public double GetLayerLimitedExpectedValue(ILayerPlus layer, Dictionary<int, IParetoParameters> curveParametersSet)
        {
            var sortedCurveParameterSet = curveParametersSet.OrderBy(c => c.Value.Threshold).ToDictionary(k => k.Key, v => v.Value);

            var layerBottom = layer.Bottom;
            var layerTop = layer.Top;

            var curveIdLessThanOrEqualToBottom = FindCurveIdLessThanOrEqualToValue(sortedCurveParameterSet, layerBottom);
            var curveLessThanOrEqualToBottom = sortedCurveParameterSet[curveIdLessThanOrEqualToBottom];

            var curveIdLessThanOrEqualToTop = FindCurveIdLessThanOrEqualToValue(sortedCurveParameterSet, layerTop);
            var curveLessThanOrEqualToTop = sortedCurveParameterSet[curveIdLessThanOrEqualToTop];

            var sortedCurvesLessThanOrEqualToTop = FindCurvesLessThanOrEqualToValue(sortedCurveParameterSet, layerTop);

            var survivalValues = PiecewiseParetoHelper.GetSurvivalValues(sortedCurvesLessThanOrEqualToTop);

            var currentCurve = sortedCurveParameterSet.First();
            var firstThreshold = currentCurve.Value.Threshold;
            if (layerTop <= firstThreshold) return layer.Limit;

            var expectedLayerLoss = 0d;
            var adjustedEffectiveBottom = layerBottom;
            var effectiveLayer = new ExcessLayer(layer.Limit, layer.Attachment);
            if (layer.Attachment < firstThreshold)
            {
                adjustedEffectiveBottom = firstThreshold;
                expectedLayerLoss = adjustedEffectiveBottom - layer.Attachment;
                effectiveLayer = new ExcessLayer(layerTop - adjustedEffectiveBottom, adjustedEffectiveBottom);
            }

            var frequency = survivalValues[curveIdLessThanOrEqualToBottom] *
                            Math.Pow(curveLessThanOrEqualToBottom.Threshold / adjustedEffectiveBottom,
                                curveLessThanOrEqualToBottom.Alpha);

            if (curveIdLessThanOrEqualToBottom == curveIdLessThanOrEqualToTop)
            {
                var lev = PiecewiseParetoHelper.GetExpectedLayerLoss(effectiveLayer, curveLessThanOrEqualToBottom.Alpha);
                expectedLayerLoss += frequency * lev;
                return expectedLayerLoss;
            }
            
            currentCurve = sortedCurveParameterSet.First(c => c.Key > curveIdLessThanOrEqualToBottom);
            var preLoopLimit = currentCurve.Value.Threshold - effectiveLayer.Attachment;
            var preLoopAttachment = effectiveLayer.Attachment;
            var preLoopLayer = new ExcessLayer(preLoopLimit, preLoopAttachment);
            var preLoopLev = PiecewiseParetoHelper.GetExpectedLayerLoss(preLoopLayer, curveLessThanOrEqualToBottom.Alpha);
            expectedLayerLoss += frequency * preLoopLev;

            var curveSubset = sortedCurveParameterSet.Where(curve => curve.Value.Threshold > curveLessThanOrEqualToBottom.Threshold &&
                                                                     curve.Value.Threshold < curveLessThanOrEqualToTop.Threshold);

            var loopExpectedLayerLoss = 0d;
            foreach (var curve in curveSubset)
            {
                currentCurve = sortedCurveParameterSet.First(c => c.Value.Threshold > curve.Value.Threshold);
                var loopLimit = currentCurve.Value.Threshold - curve.Value.Threshold;
                var loopAttachment = curve.Value.Threshold;
                var loopLayer = new ExcessLayer(loopLimit, loopAttachment);
                loopExpectedLayerLoss += survivalValues[curve.Key] *
                                         PiecewiseParetoHelper.GetExpectedLayerLoss(loopLayer, curve.Value.Alpha);
            }
            expectedLayerLoss += loopExpectedLayerLoss;

            var postLoopAttachment = sortedCurveParameterSet[curveIdLessThanOrEqualToTop].Threshold;
            var postLoopLimit = layerTop - postLoopAttachment;
            var postLoopLayer = new ExcessLayer(postLoopLimit, postLoopAttachment);

            expectedLayerLoss += survivalValues[curveIdLessThanOrEqualToTop] *
                                 PiecewiseParetoHelper.GetExpectedLayerLoss(postLoopLayer, sortedCurveParameterSet[curveIdLessThanOrEqualToTop].Alpha);

            return expectedLayerLoss;
        }

        public double GetCdfForLimit(double limit, Dictionary<int, IParetoParameters> curveParametersSet)
        {
            var sortedCurveParameterSet = curveParametersSet.OrderBy(k => k.Value.Threshold).ToDictionary(k => k.Key, v => v.Value);
            return GetCdfForLimitAlreadySorted(limit, sortedCurveParameterSet);
        }

        public static int FindCurveIdLessThanOrEqualToValue(Dictionary<int, IParetoParameters> curveParametersSet, double value)
        {
            var curvesLessThanOrEqualValue = curveParametersSet.Where(curve => curve.Value.Threshold <= value).ToList();
            if (!curvesLessThanOrEqualValue.Any()) return curveParametersSet.First().Key;

            return curvesLessThanOrEqualValue.Last().Key;
        }

        public static Dictionary<int, IParetoParameters> FindCurvesLessThanOrEqualToValue(Dictionary<int, IParetoParameters> curveParametersSet, double value)
        {
            return curveParametersSet.Where(c => c.Value.Threshold <= value).ToDictionary(k => k.Key, v => v.Value);
        }

        public Discretization Discretize(IDiscretizeInput discretizeInput, Dictionary<int, IParetoParameters> curveParametersSet)
        {
            var count = discretizeInput.BucketCount;
            var width = discretizeInput.BucketWidth;
            
            var buckets = new double[count + 2];
            for (var i = 0; i < count + 2; i++)
            {
                buckets[i] = i * width;
            }

            var sortedCurveParameterSet = curveParametersSet.OrderBy(c => c.Value.Threshold).ToDictionary(k => k.Key, v => v.Value);

            var discretization = new Discretization();
            var ns = Enumerable.Range(0, count + 1);
            var sw = Stopwatch.StartNew();
            foreach (var n in ns)
            {
                var cdfItem = CreateCdfItem(n, buckets[n], width, sortedCurveParameterSet);
                discretization.Add(new DiscretizationItem {Loss = cdfItem.Bucket, Probability = cdfItem.Cdf});
            }
            var elapsedTime = sw.ElapsedMilliseconds / 1000d;
            if (elapsedTime > 0.1) { Debug.WriteLine($"{elapsedTime:N3}");
            }


            discretization[count].Probability = 1 - discretization[count - 1].Probability;
            for (var i = discretizeInput.BucketCount - 1; i > 0; i--)
            {
                discretization[i].Probability =
                    discretization[i].Probability - discretization[i - 1].Probability;
            }

            return discretization;
        }

        internal double SimpleSecondMoment(ILayer layer, double alpha)
        {
            double secondMoment;
            if (layer.Limit.IsEqualTo(double.PositiveInfinity))
            {
                if (alpha.IsEqualTo(2))
                {
                    secondMoment = double.PositiveInfinity;
                }
                else
                {
                    secondMoment = 2 * Math.Pow(layer.Attachment, 2) * (1 / (alpha - 2) - 1 / (alpha - 1));
                }
            }
            else
            {
                if (alpha.IsEqualToZero())
                {
                    secondMoment = Math.Pow(layer.Limit, 2);
                }
                else if (alpha.IsEqualTo(1))
                {
                    secondMoment = 2 * Math.Pow(layer.Attachment, 2) *
                             (layer.Limit / layer.Attachment - Math.Log(1 + layer.Limit / layer.Attachment));
                }
                else if (alpha.IsEqualTo(2))
                {
                    secondMoment = 2 * Math.Pow(layer.Attachment, 2) *
                             ((-layer.Limit / (layer.Limit + layer.Attachment)) -
                              Math.Log(1 + layer.Limit / layer.Attachment));
                }
                else
                {
                    secondMoment = 2 * Math.Pow(layer.Attachment, 2) *
                             ((Math.Pow((1 + layer.Limit / layer.Attachment), 2 - alpha) - 1) / (2 - alpha) -
                              (Math.Pow((1 + layer.Limit / layer.Attachment), 1 - alpha) - 1) / (1 - alpha));

                }
            }

            return secondMoment;

        }

        private CdfItem CreateCdfItem(int n, double bucket, double bucketWidth, Dictionary<int, IParetoParameters> pars)
        {
            var cdf = GetCdfForLimit(bucket + bucketWidth * 0.5, pars);
            return new CdfItem { Row = n, Bucket = bucket, Cdf =  cdf};
        }

        private static double GetCdfForLimitAlreadySorted(double limit, Dictionary<int, IParetoParameters> curveParametersSet)
        {
            var sortedCurvesLessThanOrEqualToTop = FindCurvesLessThanOrEqualToValue(curveParametersSet, limit);
            if (!sortedCurvesLessThanOrEqualToTop.Any()) return 0;
            
            var survivalValues = PiecewiseParetoHelper.GetSurvivalValues(sortedCurvesLessThanOrEqualToTop);
            

            var svLast = Math.Pow(sortedCurvesLessThanOrEqualToTop.Last().Value.Threshold / limit,
                sortedCurvesLessThanOrEqualToTop.Last().Value.Alpha);
            
            return 1 - svLast * survivalValues.Last().Value; //1-(svLast * prod);
        }

        private class CdfItem
        {
            public int Row { get; set; }
            public double Bucket { get; set; }
            public double Cdf { get; set; }
        }

    }
}
