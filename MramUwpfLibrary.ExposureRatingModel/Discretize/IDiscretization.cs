using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Layers;
using MramUwpfLibrary.Common.Policies;

namespace MramUwpfLibrary.ExposureRatingModel.Discretize
{
    public interface IDiscretization : IList<IDiscretizationItem>,ICloneable
    {
        IDiscretization BalanceBackToExpectedLoss(ITower tower);
        IDiscretization OverlayPolicySet(IPolicySet policySet);
        double ExpectedLoss { get; }
        double VarianceOfLoss { get; }
    }

    public class Discretization : IDiscretization
    {
        private readonly IList<IDiscretizationItem> _discretizationItems;
        private const double Threshold = 0;
        
        public Discretization()
        {
            _discretizationItems = new List<IDiscretizationItem>();
        }

        public IEnumerable<double> GetMassPoints()
        {
            GroupByAndSum();
            var nodes = CreateNodes();

            var massPoints = new List<double>();
            var currentNode = nodes.First;
            while (currentNode != null)
            {
                var previousNode = currentNode.Previous;
                if (currentNode.Value.Probability / previousNode?.Value.Probability - 1 > Threshold)
                {
                    massPoints.Add(currentNode.Value.Loss);
                }

                currentNode = currentNode.Next;
            }

            return massPoints;
        }

        public IDiscretization Scale(IPolicy policy)
        {
            var discretization = new Discretization();

            var topCounter = 0;
            var bottomCounter = 0;
            var policyBottom = ((Policy) policy).Bottom;
            var policyTop = ((Policy) policy).Top;

            var probabilityBelowBottom = this.Where(item => item.Loss < policyBottom).Sum(item => item.Probability);
            var probabilityAboveTop = this.Where(item => item.Loss > policyTop).Sum(item => item.Probability);
            foreach (var item in this)
            {
                var newItem = new DiscretizationItem { Loss = item.Loss };

                if (item.Loss < policyBottom) bottomCounter++;
                if (item.Loss < policyTop) topCounter++;
                if (item.Loss >= policyBottom && item.Loss <= policyTop) newItem.Probability = item.Probability;

                discretization.Add(newItem);
            }
            
            discretization[bottomCounter].Probability += probabilityBelowBottom;
            discretization[topCounter].Probability += probabilityAboveTop;

            return discretization;
        }

        public IDiscretization BalanceBackToExpectedLoss(ITower tower)
        {
            var firstLoss = tower.First().Value.Loss;

            var targetSigmaGx = tower
                .Select(layerLossEstimate => new LayerLossEstimate(layerLossEstimate.Value, layerLossEstimate.Value.Loss / firstLoss))
                .ToList();

            var gX = SurvivalDistribution.GetSurvivalDistribution(this);

            var currentSigmaGx = new Tower();
            var scaleFactors = new Tower();
            foreach (var layer in targetSigmaGx)
            {
                currentSigmaGx.Add(new LayerLossEstimate(layer, gX.SumBand(layer.Limit, layer.Attachment)));
                scaleFactors.Add(new LayerLossEstimate(layer,
                    layer.Loss / (currentSigmaGx.Last().Value.Loss / currentSigmaGx.First().Value.Loss)));
            }

            foreach (var item in gX.Items)
            {
                foreach (var scaleFactor in scaleFactors)
                {
                    if (item.Loss > scaleFactor.Value.Top) continue;
                    item.Probability *= scaleFactor.Value.Loss;
                    break;
                }
            }

            return gX.MapToDiscretization();
        }

        #region list support

        public IEnumerator<IDiscretizationItem> GetEnumerator()
        {
            return _discretizationItems.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(IDiscretizationItem item)
        {
            _discretizationItems.Add(item);
        }

        public void Clear()
        {
            _discretizationItems.Clear();
        }

        public bool Contains(IDiscretizationItem item)
        {
            return _discretizationItems.Contains(item);
        }

        public void CopyTo(IDiscretizationItem[] array, int arrayIndex)
        {
            _discretizationItems.CopyTo(array, arrayIndex);
        }

        public bool Remove(IDiscretizationItem item)
        {
            return _discretizationItems.Remove(item);
        }

        public int Count => _discretizationItems.Count;
        public bool IsReadOnly => false;

        public int IndexOf(IDiscretizationItem item)
        {
            return _discretizationItems.IndexOf(item);
        }

        public void Insert(int index, IDiscretizationItem item)
        {
            _discretizationItems.Insert(index, item);
        }

        public void RemoveAt(int index)
        {
            _discretizationItems.RemoveAt(index);
        }

        public IDiscretizationItem this[int index]
        {
            get => _discretizationItems[index];
            set => _discretizationItems[index] = value;
        }

        #endregion

        public object Clone()
        {
            var discretization = new Discretization();
            foreach (var item in this)
            {
                discretization.Add((DiscretizationItem) item.Clone());
            }
            return discretization;
        }

        public IDiscretization OverlayPolicy(IPolicy policy)
        {
            var policySir = policy.Sir.GetValueOrDefault();
            var policyBottom = ((Policy) policy).Bottom;
            var policyTop = ((Policy) policy).Top;

            var discretization = (Discretization) Clone();

            if (double.IsPositiveInfinity(policy.Limit) && policy.Sir.GetValueOrDefault().IsEqualToZero()) return discretization;

            if (policyBottom.IsEpsilonEqualToZero())
            {
                foreach (var item in discretization.Where(item => item.Loss > policy.Limit))
                {
                    item.Probability = 0;
                }
            }
            else
            {
                foreach (var item in discretization)
                {
                    item.Probability = 0;
                }

                var sirRow = this.SingleOrDefault(item => item.Loss.IsEpsilonEqual(policySir));
                if (sirRow != null)
                {
                    discretization.First().Probability = this.Where(item => item.Loss <= policySir).Sum(d => d.Probability);
                }
                else
                {
                    discretization.First().Probability = this.Where(d => d.Loss < policySir).Sum(d => d.Probability)
                                                         + this.First(d => d.Loss > policySir).Probability;
                }

                var sirShift = this.Count(item => item.Loss < policySir);
                var exposedCount = this.Count(item => item.Loss < policy.Limit);

                for (var i = 1; i < exposedCount; i++)
                {
                    discretization[i].Probability = this[i + sirShift].Probability;
                }
            }

            if (!this.Any(item => item.Loss > policyTop)) return discretization;

            var sum = this.Where(item => item.Loss >= policyTop).Sum(item => item.Probability);
            var count = discretization.Count(item => item.Loss < policy.Limit);
            discretization[count].Probability = sum;
            return discretization;
        }

        public IDiscretization OverlayPolicySet(IPolicySet policySet)
        {
            var discretizations = policySet.Select(policy => (IDiscretizationSet) new DiscretizationSet
            {
                Weight = policy.Amount,
                Discretization = OverlayPolicy(policy)
            }).ToList();

            var discretizationNetOfPolicies = DiscretizationSet.Aggregate(discretizations);
            return discretizationNetOfPolicies;
        }

        public static IDiscretizeInput GetDiscretizeInput(double layerTop, bool isLayerWithRespectToGround, IPolicySet policySet)
        {
            const int bucketCountMinusOne = CollectiveModelsConstants.BucketCount - 1;

            var maximumPolicyTop = policySet.Where(policy => !double.IsPositiveInfinity(policy.Limit))
                .Max(policy => policy.Limit + policy.Sir.GetValueOrDefault());

            var maximumPolicyAttachment = policySet.Max(policy => policy.Sir.GetValueOrDefault());
            var maximumLayerTop = isLayerWithRespectToGround ? layerTop : layerTop + maximumPolicyAttachment;

            var bucketWidth = (int) Math.Ceiling(Math.Max(maximumPolicyTop, maximumLayerTop) / bucketCountMinusOne);
            return new DiscretizeInput(CollectiveModelsConstants.BucketCount, bucketWidth);
        }

        public static IList<DiscretizationItem> Reduce(IList<IDiscretizationItem> items)
        {
            return items
                .GroupBy(item => item.Loss, new DoubleComparer())
                .Select(group => new DiscretizationItem
                {
                    Loss = group.Key,
                    Probability = group.Sum(s => s.Probability)
                }).ToList();
        }

        private void GroupByAndSum()
        {
            var distinctCount = _discretizationItems.Select(item => item.Loss).Distinct(new DoubleComparer()).Count();
            if (distinctCount == _discretizationItems.Count) return;

            var reducedItems = Reduce(_discretizationItems);

            _discretizationItems.Clear();
            foreach (var item in reducedItems)
            {
                _discretizationItems.Add(new DiscretizationItem {Loss = item.Loss, Probability = item.Probability});
            }
        }

        private LinkedList<IDiscretizationItem> CreateNodes()
        {
            var sortedList = new SortedList<double, IDiscretizationItem>();
            foreach (var item in this)
            {
                sortedList.Add(item.Loss, item);
            }
            
            var nodes = new LinkedList<IDiscretizationItem>();
            foreach (var item in sortedList)
            {
                nodes.AddLast(item.Value);
            }
            return nodes;
        }

        public IDiscretization MapCumulativeToIncremental()
        {
            var discretization = new Discretization
            {
                new DiscretizationItem
                {
                    Loss = _discretizationItems.First().Loss, Probability = _discretizationItems.First().Probability
                }
            };

            for (var i = 1; i < Count; i++)
            {
                discretization.Add(new DiscretizationItem
                {
                    Loss = _discretizationItems[i].Loss,
                    Probability = _discretizationItems[i].Probability - _discretizationItems[i - 1].Probability
                });
            }

            return discretization;
        }

        public double ExpectedLoss => _discretizationItems.Select(item => item.Loss * item.Probability).Sum();
        public double ExpectedSquaredLoss => _discretizationItems.Select(item => item.Loss.Squared() * item.Probability).Sum();
        public double VarianceOfLoss => ExpectedSquaredLoss - ExpectedLoss.Squared();
    }
}