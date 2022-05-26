using System.Collections.Generic;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;
// ReSharper disable PossibleInvalidOperationException

namespace PionlearClient.Model
{
    public class TotalInsuredValueModel : BaseSourceComponentModel
    {
        public IList<TotalInsuredValueDistributionItemPlus> Items { get; set; }

        public override string FriendlyName => BexConstants.TotalInsuredValueProfileName;
        public int? UmbrellaTypeId { get; set; }
        
        public override StringBuilder Validate(IList<CollectorApi.Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();

            foreach (var item in Items)
            {
                var location = item.Location;

                if (double.IsNaN(item.TotalInsuredValue) || item.TotalInsuredValue <= 0)
                {
                    messages.AppendLine($"Change {BexConstants.TivName} <{item.TotalInsuredValue:N0}> to a positive number in {location}");
                }

                if (item.Share.HasValue && double.IsNaN(item.Share.Value) || item.Share < 0 || item.Share > 1)
                {
                    messages.AppendLine($"Change {BexConstants.ShareName.ToLower()} <{item.Share:N0}> " +
                                          $"to a number between 0 and 1 in {location}");
                }

                if (item.Limit.HasValue && double.IsNaN(item.Limit.Value) || item.Limit < 0)
                {
                    messages.AppendLine($"Change {BexConstants.LimitName.ToLower()} <{item.Limit:N0}> " +
                                          $"to a non-negative number in {location}");
                }

                if (item.Attachment.HasValue && double.IsNaN(item.Attachment.Value) || item.Attachment < 0)
                {
                    messages.AppendLine($"Change {BexConstants.SirAttachmentName.ToLower()} <{item.Attachment:N0}> " +
                                          $"to a non-negative number in {location}");
                }

                if (double.IsNaN(item.Weight) || item.Weight < 0 || item.Weight > 1)
                {
                    messages.AppendLine($"Change allocation value <{item.Weight:P2}> to a number between 0 and 1 in {location}");
                }
            }

            return messages;
        }

        public TotalInsuredValueDistributionModel Map()
        {
            return new TotalInsuredValueDistributionModel
            {
                Id = SourceId,
                DistributionName = Name,
                SublineIds = SublineIds.Select(id => id.Value).ToList(),
                Items = Items.Select(item => new TotalInsuredValueAndWeight()
                {
                    TotalInsuredValue = item.TotalInsuredValue,
                    Limit = item.Limit,
                    Attachment = item.Attachment,
                    Share = item.Share,
                    Weight = item.Weight
                }).ToList()
            };
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
    }

    internal static class TotalInsuredValuesExtensions
    {
        public static void ValidateEmptiness(this IList<TotalInsuredValueModel> models, StringBuilder validation)
        {
            var nonEmptyCount = models.Count(model => !model.Items.Any());
            var emptyCount = models.Count(model => model.Items.Any());
            if (emptyCount > 0 && nonEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.TotalInsuredValueProfileName.ToLower()} contains data");
            }
        }

        public static void ValidateSublineComposition(this IList<TotalInsuredValueModel> models, StringBuilder validation, long sublineId, string name)
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
                    if (subsetModels.Count != 1 && !subsetModels.All(model => model.UmbrellaTypeId.HasValue))
                    {
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.TotalInsuredValueProfileName.ToLower()}");
                    }
                }
                else
                {
                    validation.AppendLine($"{name} doesn't appear in a {BexConstants.TotalInsuredValueProfileName.ToLower()}");
                }
            }
            else
            {
                if (subsetModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.TotalInsuredValueProfileName.ToLower()}");
            }


        }
    }
}