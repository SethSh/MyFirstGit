using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using WorkersCompStateClassCodeAndValue = MunichRe.Bex.ApiClient.CollectorApi.WorkersCompStateClassCodeAndValue;

namespace PionlearClient.Model
{
    public class WorkersCompClassCodeModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.WorkersCompClassCodeProfileName;
        public IList<WorkersCompStateClassCodeAndValuePlus> Items { get; set; }

        
        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();
            CheckAggregatedNegatives(messages);
            return messages;
        }

        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            return messages;
        }

        public WorkersCompStateClassCodeDistributionModel Map()
        {
            //removing items without hazard group id
            return new WorkersCompStateClassCodeDistributionModel
            {
                Id = SourceId,
                SublineIds = SublineIds.Select(id =>
                {
                    Debug.Assert(id != null, nameof(id) + " != null");
                    return id.Value;
                }).ToList(),
                Items = Items
                    .Where(item => item.HasHazard)
                    .Select(item => new WorkersCompStateClassCodeAndValue
                {
                    WorkersCompStateId = item.WorkersCompStateId,
                    ClassCodeId = item.ClassCodeId,
                    Value = item.Value
                }).ToList()
            };
        }

        private void CheckAggregatedNegatives(StringBuilder messages)
        {
            const int displayMaximumLength = 10;
            const int displayMaximumRowLength = 10;

            var aggregatedNegativeItems = Items
                .GroupBy(item => new { item.WorkersCompStateId , item.ClassCodeId })
                .Select(item => new WorkersCompStateClassCodeAndValue
                {
                    WorkersCompStateId = item.First().WorkersCompStateId,
                    ClassCodeId = item.First().ClassCodeId,
                    Value = item.Sum(c => c.Value),
                })
                .Where(item => item.Value < 0)
                .ToList();
            
            if (!aggregatedNegativeItems.Any()) return; 
            
            var stateDictionary = StateCodesFromBex.GetWorkersCompStates().ToDictionary(key => Convert.ToInt64(key.Id), value => value.Abbreviation);
            var stateIds = aggregatedNegativeItems.Select(item => item.WorkersCompStateId);
            var classCodeByStateDictionary = WorkersCompClassCodesAndHazardsFromBex.GetClassCodeByStateDictionary(stateIds);
            
            foreach (var aggregatedItem in aggregatedNegativeItems.Take(displayMaximumLength))
            {
                var matches = Items.Where(item => item.WorkersCompStateId == aggregatedItem.WorkersCompStateId && item.ClassCodeId == aggregatedItem.ClassCodeId);
                var rowNumbers = matches.Select(item => item.RowNumber).ToList();
                var rowCount = rowNumbers.Count;

                var stateAbbreviation = stateDictionary[aggregatedItem.WorkersCompStateId];
                var classCodeDictionary = classCodeByStateDictionary[aggregatedItem.WorkersCompStateId];
                var classCode = classCodeDictionary[aggregatedItem.ClassCodeId].StateClassCode.ToString("0000");

                var message = rowNumbers.Count == 1
                    ? $"<{stateAbbreviation} {classCode}> contains negative premium <{aggregatedItem.Value:N0}> in row: "
                    : $"<{stateAbbreviation} {classCode}> contains negative premium sum <{aggregatedItem.Value:N0}> in rows: ";
                
                if (rowCount > displayMaximumRowLength)
                {
                    message += string.Join(", ", rowNumbers.Take(displayMaximumRowLength).ToArray()) + " ...";
                }
                else
                {
                    message += string.Join(", ", rowNumbers.ToArray());
                }

                messages.AppendLine(message);
            }

            if (aggregatedNegativeItems.Count <= displayMaximumLength) return;

            var notDisplayedCount = aggregatedNegativeItems.Count - displayMaximumLength;
            messages.AppendLine(
                notDisplayedCount == 1
                    ? $"There is {notDisplayedCount:N0} additional negative not displayed"
                    : $"There are {notDisplayedCount:N0} additional negatives not displayed");
        }
    }

    internal static class WorkerCompClassCodeExtensions
    {
        public static void ValidateSublineComposition(this IList<WorkersCompClassCodeModel> models, StringBuilder validation, long sublineId, string name)
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
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.WorkersCompClassCodeProfileName.ToLower()}");
                }
                else
                {
                    validation.AppendLine(
                        $"{name} doesn't appear in a {BexConstants.WorkersCompClassCodeProfileName.ToLower()}");
                }
            }
            else
            {
                if (subsetModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.WorkersCompClassCodeProfileName.ToLower()}");
            }
        }
    }
}