using System.Collections.Generic;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using PionlearClient.Extensions;
// ReSharper disable PossibleInvalidOperationException

namespace PionlearClient.Model
{
    public class StateModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.StateProfileName;
        public IList<StateDistributionItemPlus> Items { get; set; }
        
        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();

            #region not valid when contains auto liability personal and only CW
            if (riskClassValidSublines.Select(subline => subline.Id).Contains(BexConstants.AutoLiabilityPersonalSublineCode))
            {
                if (Items.Count == 1)
                {
                    var onlyItem = Items.First();
                    if (onlyItem.Value > 0 && onlyItem.StateCode == BexConstants.StateProfileCountrywideId)
                    {
                        var subline = SublineCodesFromBex.ReferenceData.First(x =>
                            x.SublineId == BexConstants.AutoLiabilityPersonalSublineCode);
                        var sublineFriendly = subline.LineOfBusiness.Name.ConnectWithDash(subline.SublineName);

                        var stateName = StateCodesFromBex.ReferenceData.First(s => s.Id == BexConstants.StateProfileCountrywideId).Name;
                        messages.AppendLine($"Can't contain subline <{sublineFriendly}> and only the single state <{stateName}>");
                    }
                }
            }
            #endregion

            foreach (var item in Items)
            {
                if (item.Value < 0 || item.Value > 1)
                {
                    messages.AppendLine($"Change allocation value <{item.Value:P2}> " +
                                          $"to a number between 0 and 1 in {item.Location}");
                }

            }

            return messages;
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

        public StateDistributionModel Map()
        {
            return new StateDistributionModel
            {
                Id = SourceId,
                DistributionName = Name,
                SublineIds = SublineIds.Select(id => id.Value).ToList(),
                Items = Items.Select(item => new StateDistributionItem
                {
                    StateCode = item.StateCode,
                    Value = item.Value
                }).ToList()
            };
        }
    }

    internal static class StateModelsExtensions
    {
        public static void ValidateEmptiness(this IList<StateModel> models, StringBuilder validation)
        {
            var nonEmptyCount = models.Count(model => !model.Items.Any());
            var emptyCount = models.Count(model => model.Items.Any());
            if (emptyCount > 0 && nonEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.StateProfileName.ToLower()} contains data");
            }
        }

        public static void ValidateSublineComposition(this IList<StateModel> models, StringBuilder validation, long sublineId, string name)
        {
            var subsetModels = models.Where(x => x.SublineIds.Contains(sublineId)).ToList();
            var hasProfile = !SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == sublineId).HasDefaultState;
            if (hasProfile)
            {
                if (subsetModels.Any())
                {
                    if (subsetModels.Count != 1)
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.StateProfileName.ToLower()}");
                }
                else
                {
                    validation.AppendLine(
                        $"{name} doesn't appear in a {BexConstants.StateProfileName.ToLower()}");
                }
            }
            else
            {
                if (subsetModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.StateProfileName.ToLower()}");
            }
        }
    }
}