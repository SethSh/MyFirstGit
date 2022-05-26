using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;

namespace PionlearClient.Model
{
    public class WorkersCompStateAttachmentModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.WorkersCompStateAttachmentProfileName;
        public IList<WorkersCompStateAttachmentAndWeightPlus> Items { get; set; }

        
        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();

            var duplicateAttachments = GetDuplicateAttachments();
            if (duplicateAttachments.Any())
            {
                var values = string.Join(" ", duplicateAttachments.Select(d => d.ToString("N0")).ToArray());
                messages.AppendLine($"Remove the duplicate attachment(s) <{values}>");
            }

            var statesWithTooMuchPercent = GetStatesWithTooMuchPercent();
            if (statesWithTooMuchPercent.Any())
            {
                var values = string.Join(", ", statesWithTooMuchPercent.ToArray());
                messages.AppendLine($"These state percent(s) add to greater then {1:P0} <{values}>");
            }

            foreach (var item in Items)
            {
                if (item.Attachment <= 0)
                {
                    messages.AppendLine($"Change attachment value <{item.Weight:P2}> " +
                                        $"to a positive number in {item.Location}");
                }

                if (item.Weight < 0 )
                {
                    messages.AppendLine($"Change premium value <{item.Weight:P2}> " +
                                          $"to a non-negative number in {item.Location}");
                }
            }


            return messages;
        }

        private IList<string> GetStatesWithTooMuchPercent()
        {
            //looking for bad inputs
            //close to 100% is fine
            var list = new List<string>();
            const double eps = 0.005;
            var stateIds = Items.Select(item => item.WorkersCompStateId).Distinct();

            var states = StateCodesFromBex.GetWorkersCompStates().ToDictionary(state => state.Id);

            foreach (var stateId in stateIds)
            {
                var statePercentTotal = Items.Where(item => item.WorkersCompStateId == stateId).Sum(item => item.Weight);
                if (statePercentTotal > 1d + eps)
                {
                    list.Add(states[Convert.ToInt32(stateId)].Abbreviation);
                }
            }

            return list;
        }

        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            return messages;
        }

        public WorkersCompStateAttachmentDistributionModel Map()
        {
            return new WorkersCompStateAttachmentDistributionModel
            {
                Id = SourceId,
                SublineIds = SublineIds.Select(id => id.Value).ToList(),
                Items = Items.Select(item => new WorkersCompStateAttachmentAndWeight
                {
                    WorkersCompStateId= item.WorkersCompStateId,
                    Attachment = item.Attachment,
                    Weight = item.Weight
                }).ToList()
            };
        }

        public IList<double> GetDuplicateAttachments()
        {
            var list = new List<double>();
            var attachments = Items.Select(item => item.Attachment).Distinct();
            foreach (var attachment in attachments)
            {
                var attachmentMatches = Items.Where(item => item.Attachment.IsEpsilonEqual(attachment)).ToList();
                //if a state appears more than once then mark the attachment as a duplicate
                if (attachmentMatches.Select(s => s.WorkersCompStateId).Distinct().Count() < attachmentMatches.Select(s => s.WorkersCompStateId).Count())
                {
                    list.Add(attachment);
                }
            }

            return list;
        }
    }

    internal static class WorkerCompStateAttachmentExtensions
    {
        public static void ValidateSublineComposition(this IList<WorkersCompStateAttachmentModel> models, StringBuilder validation, long sublineId, string name)
        {
            //retrofitting old workbooks that may not have any models
            if (!models.Any()) return; 
            
            var subsetModels = models.Where(model => model.SublineIds.Contains(sublineId)).ToList();
            //todo get this into table and out of code
            var hasProfile = SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == sublineId).LineOfBusiness.Name == "Workers Compensation";

            if (hasProfile)
            {
                if (subsetModels.Any())
                {
                    if (subsetModels.Count != 1)
                        validation.AppendLine($"{name} is contained in more than one {BexConstants.WorkersCompStateAttachmentProfileName.ToLower()}");
                }
                else
                {
                    validation.AppendLine(
                        $"{name} doesn't appear in a {BexConstants.WorkersCompStateAttachmentProfileName.ToLower()}");
                }
            }
            else
            {
                if (subsetModels.Any())
                    validation.AppendLine(
                        $"{name} can't be contained in a {BexConstants.WorkersCompStateAttachmentProfileName.ToLower()}");
            }
        }
    }
}