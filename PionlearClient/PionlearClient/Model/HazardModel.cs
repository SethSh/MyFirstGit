using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using MunichRe.Bex.ApiClient.CollectorApi;
// ReSharper disable PossibleInvalidOperationException

namespace PionlearClient.Model
{
    public class HazardModel : BaseSourceComponentModel
    {
        public IList<HazardDistributionItemPlus> Items { get; set; }

        public override string FriendlyName => BexConstants.HazardProfileName;
        
        public int SublineCode { get; set; }
        
        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();
            foreach (var item in Items)
            {
                var location = item.Location;
                if (double.IsNaN(item.Value))
                {
                    messages.AppendLine($"Change allocation to a number in {location}"); 
                }
                else
                {
                    if (item.Value < 0 || item.Value > 1)
                    {
                        messages.AppendLine($"Change allocation value <{item.Value:P2}> to a number between 0 and 1 in {location}");
                    }
                }
            }
            return messages;
        }

        public HazardDistributionModel Map()
        {
            return new HazardDistributionModel
            { 
                Id = SourceId,
                DistributionName = Name,
                SublineIds = SublineIds.Select(id => id.Value).ToList(),
                Items = Items.Select(item => new HazardDistributionItem
                {
                    HazardId = Convert.ToInt16(item.HazardId),
                    Value = item.Value
                }).ToList()
            };
        }
        
        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();

            const double tolerance = NumericalConstants.QcProfileTolerance;
            var valueSum = Items.Sum(item => item.Value);
            if (!valueSum.IsEpsilonEqual(1, tolerance))
            {
                messages.AppendLine($"The percent sum <{valueSum:P4}> is not within {tolerance:P4} of {1:P4}");
            }
            return messages;
        }
    }

    internal static class HazardModelsExtension
    {
        public static void ValidateEmptiness(this IList<HazardModel> models, StringBuilder validation)
        {
            var nonEmptyCount = models.Count(model => !model.Items.Any());
            var emptyCount = models.Count(model => model.Items.Any());
            if (emptyCount > 0 && nonEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.HazardProfileName.ToLower()} contains data");
            }
        }
        
        public static void ValidateSublineComposition(this IList<HazardModel> models, StringBuilder validation, long sublineId, string name)
        {
            //temporary until wc allows for exposure rating
            var sublineIdsThatDoNotExposureRate = GetSublineIdsThatDoNotExposureRate();

            var subsetModels = models.Where(x => x.SublineIds.Contains(sublineId)).ToList();

            bool hasProfile;
            if (sublineIdsThatDoNotExposureRate.Contains(sublineId))
            {
                hasProfile = false;
            }
            else
            {
                var hasDefaultProfile = SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == sublineId).HasDefaultHazard;
                hasProfile = !hasDefaultProfile;
            }

            if (hasProfile)
            {
                if (subsetModels.Any())
                {
                    if (subsetModels.Count != 1)
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.HazardProfileName.ToLower()}");
                }
                else
                {
                    validation.AppendLine(
                        $"{name} doesn't appear in a {BexConstants.HazardProfileName.ToLower()}");
                }
            }
            else
            {
                if (subsetModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.HazardProfileName.ToLower()}");
            }
        }

        private static IList<long> GetSublineIdsThatDoNotExposureRate()
        {
            var propertySublineIds = SublineCodesFromBex.GetPropertySublines().Select(subline => subline.SublineId).ToList();
            var autoPhysicalDamageSublineIds =
                SublineCodesFromBex.GetAutoPhysicalDamageSublines().Select(subline => subline.SublineId).ToList();
            var workersCompensationSublineIds =
                SublineCodesFromBex.GetWorkersCompensationSublineIds().Select(subline => subline.SublineId).ToList();

            var sublineIdsThatDoNotExposureRate = new List<long>();
            foreach (var id in propertySublineIds)
            {
                sublineIdsThatDoNotExposureRate.Add(Convert.ToInt32(id));
            }

            foreach (var id in autoPhysicalDamageSublineIds)
            {
                sublineIdsThatDoNotExposureRate.Add(Convert.ToInt32(id));
            }

            foreach (var id in workersCompensationSublineIds)
            {
                sublineIdsThatDoNotExposureRate.Add(Convert.ToInt32(id));
            }

            return sublineIdsThatDoNotExposureRate;
        }
    }
}