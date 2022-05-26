using System.Collections.Generic;
using System.Linq;
using System.Text;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using MunichRe.Bex.ApiClient.CollectorApi;
// ReSharper disable PossibleInvalidOperationException

namespace PionlearClient.Model
{
    public class PolicyModel : BaseSourceComponentModel
    {
        public IList<PolicyDistributionItemPlus> Items { get; set; }

        public override string FriendlyName => BexConstants.PolicyProfileName;
        public int? UmbrellaTypeId { get; set; }
        
        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();

            foreach (var item in Items)
            {
                var location = item.Location;

                if (double.IsNaN(item.Limit) || item.Limit <= 0)
                {
                    messages.AppendLine($"Change {BexConstants.LimitName.ToLower()} <{item.Limit:N0}> to a positive number in {location}");
                }

                if (item.Attachment.HasValue && double.IsNaN(item.Attachment.Value) || item.Attachment < 0)
                {
                    messages.AppendLine($"Change {BexConstants.SirAttachmentName.ToLower()} <{item.Attachment:N0}> " +
                                          $"to a non-negative number in {location}");
                }

                if (double.IsNaN(item.Value) || item.Value < 0 || item.Value > 1)
                {
                    messages.AppendLine($"Change allocation value <{item.Value:P2}> to a number between 0 and 1 in {location}");
                }
            }

         
            return messages;
        }

        public PolicyDistributionModel Map()
        {
            return new PolicyDistributionModel
            {
                Id = SourceId,
                DistributionName = Name,
                UmbrellaTypeId = UmbrellaTypeId,
                SublineIds = SublineIds.Select(id => id.Value).ToList(),
                Items = Items.Select(item => new PolicyDistributionItem
                {
                    Limit = item.Limit, 
                    Attachment = item.Attachment, 
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
                messages.AppendLine($"The percent sum <{valueSum:P4}> isn't within {tolerance:P4} of {1:P4}");
            }

            return messages;
        }
    }

    internal static class PolicyModelsExtensions
    {
        public static void ValidateEmptiness(this IList<PolicyModel> models, StringBuilder validation)
        {
            var nonEmptyCount = models.Count(model => !model.Items.Any());
            var emptyCount = models.Count(model => model.Items.Any());
            if (emptyCount > 0 && nonEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.PolicyProfileName.ToLower()} contains data");
            }
        }

        public static void ValidateSublineComposition(this IList<PolicyModel> models, StringBuilder validation, long sublineId, string name)
        {
            var sublineModels = models.Where(x => x.SublineIds.Contains(sublineId)).ToList();
            var hasProfile = SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == sublineId).LineOfBusiness.IncludesPolicyLimitProfile;
            if (hasProfile)
            {
                if (sublineModels.Any())
                {
                    if (sublineModels.Count != 1 && !sublineModels.All(model => model.UmbrellaTypeId.HasValue))
                    {
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.PolicyProfileName.ToLower()}");
                    }
                }
                else
                {
                    validation.AppendLine($"{name} doesn't appear in a {BexConstants.PolicyProfileName.ToLower()}");
                }
            }
            else
            {
                if (sublineModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.PolicyProfileName.ToLower()}");
            }


        }
    }
}