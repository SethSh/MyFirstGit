using System.Collections.Generic;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;

namespace PionlearClient.Model
{
    public class ProtectionClassModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.ProtectionClassProfileName;
        public IList<ProtectionClassDistributionItemPlus> Items { get; set; }

        public int SublineCode { get; set; }

        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();
            foreach (var item in Items)
            {
                if (item.Weight < 0 || item.Weight > 1)
                {
                    messages.AppendLine($"Change allocation value <{item.Weight:P2}> " +
                                          $"to a number between 0 and 1 in {item.Location}");
                }

            }
            return messages;
        }

        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();

            const double tolerance = NumericalConstants.QcProfileTolerance;
            var valueSum = Items.Sum(item => item.Weight);
            if (!valueSum.IsEpsilonEqual(1, tolerance))
            {
                messages.AppendLine($"The percent sum <{valueSum:P4}> isn't within {tolerance:P4} of {1:P4}");
            }

            return messages;
        }

        public ProtectionClassDistributionModel Map()
        {
            return new ProtectionClassDistributionModel
            {
                Id = SourceId,
                DistributionName = Name,
                // ReSharper disable once PossibleInvalidOperationException
                SublineIds = SublineIds.Select(id => id.Value).ToList(),
                Items = Items.Select(item => new ProtectionClassAndWeight()
                {
                    ProtectionClassId = item.ProtectionClassId,
                    Weight = item.Weight
                }).ToList()
            };
        }
    }

    internal static class ProtectionClassesExtensions
    {
        public static void ValidateEmptiness(this IList<ProtectionClassModel> models, StringBuilder validation)
        {
            var nonEmptyCount = models.Count(model => !model.Items.Any());
            var emptyCount = models.Count(model => model.Items.Any());
            if (emptyCount > 0 && nonEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.ProtectionClassProfileName.ToLower()} contains data");
            }
        }

        public static void ValidateSublineComposition(this IList<ProtectionClassModel> models, StringBuilder validation, long sublineId, string name)
        {
            //retrofitting old workbooks that may not have any models
            if (!models.Any()) return; 
            
            var subsetModels = models.Where(x => x.SublineIds.Contains(sublineId)).ToList();
            //todo get this into table and out of code
            var hasProfile = SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == sublineId).LineOfBusiness.Name == "Property";

            if (hasProfile)
            {
                if (subsetModels.Any())
                {
                    if (subsetModels.Count != 1)
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.ProtectionClassProfileName.ToLower()}");
                }
                else
                {
                    validation.AppendLine(
                        $"{name} doesn't appear in a {BexConstants.ProtectionClassProfileName.ToLower()}");
                }
            }
            else
            {
                if (subsetModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.ProtectionClassProfileName.ToLower()}");
            }
        }
    }
}