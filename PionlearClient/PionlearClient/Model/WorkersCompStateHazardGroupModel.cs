using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;

namespace PionlearClient.Model
{
    public class WorkersCompStateHazardGroupModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.WorkersCompStateHazardGroupProfileName;
        public IList<WorkersCompStateHazardGroupAndWeightPlus> Items { get; set; }

        
        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();
            foreach (var item in Items)
            {
                if (item.Value < 0)
                {
                    messages.AppendLine($"Change allocation value <{item.Value:N0}> " +
                                          $"to a non-negative number in {item.Location}");
                }

            }
            return messages;
        }

        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            return messages;
        }

        public WorkersCompStateHazardGroupDistributionModel Map()
        {
            return new WorkersCompStateHazardGroupDistributionModel
            {
                Id = SourceId,
                SublineIds = SublineIds.Select(id =>
                {
                    Debug.Assert(id != null, nameof(id) + " != null");
                    return id.Value;
                }).ToList(),
                Items = Items.Select(item => new WorkersCompStateHazardGroupAndWeight
                {
                    WorkersCompStateId = item.WorkersCompStateId,
                    WorkersCompHazardGroupId= item.WorkersCompHazardGroupId,
                    Value = item.Value
                }).ToList()
            };
        }
    }

    internal static class WorkerCompStateAttHazardExtensions
    {
        public static void ValidateSublineComposition(this IList<WorkersCompStateHazardGroupModel> models, StringBuilder validation, long sublineId, string name)
        {
            //retrofitting old workbooks that may not have any models
            if (!models.Any()) return; 
            
            var subsetModels = models.Where(x => x.SublineIds.Contains(sublineId)).ToList();
            //todo get this into table and out of code
            var hasProfile = SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == sublineId).LineOfBusiness.Name == "Workers Compensation";

            if (hasProfile)
            {
                if (subsetModels.Any())
                {
                    if (subsetModels.Count != 1)
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.WorkersCompStateHazardGroupProfileName.ToLower()}");
                }
                else
                {
                    validation.AppendLine(
                        $"{name} doesn't appear in a {BexConstants.WorkersCompStateHazardGroupProfileName.ToLower()}");
                }
            }
            else
            {
                if (subsetModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.WorkersCompStateHazardGroupProfileName.ToLower()}");
            }
        }
    }
}