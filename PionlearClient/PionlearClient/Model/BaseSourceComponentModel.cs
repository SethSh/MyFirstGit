using System.Collections.Generic;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;

namespace PionlearClient.Model
{
    public abstract class BaseSourceComponentModel : BaseSourceModel, ISourceComponentModel
    {
        public abstract string FriendlyName { get; }
        public abstract StringBuilder Validate(IList<Allocation> riskClassValidSublines);
        public abstract StringBuilder PerformQualityControl();

        public void BuildQualityControlMessage(StringBuilder stringBuilder,  string segmentName)
        {
            var messages = PerformQualityControl();
            if (messages.Length == 0) return;

            stringBuilder.AppendLine($"{segmentName.ConnectWithDash(Name)}");
            stringBuilder.Append(messages);
            stringBuilder.AppendLine();
        }

        public void BuildValidationMessage(IList<Allocation> sublineAllocations, StringBuilder stringBuilder, string segmentName)
        {
            var messages = Validate(sublineAllocations);

            var sublineIds = sublineAllocations.Select(x => x.Id);
            var orphanSublineIds = SublineIds.Where(id => id.HasValue).Select(id => id.Value).Except(sublineIds);
            foreach (var item in orphanSublineIds)
            {
                var subline = SublineCodesFromBex.ReferenceData.First(x => x.SublineId == item);
                var name = subline.LineOfBusiness.Name.ConnectWithDash(subline.SublineName);
                var butNot = $"but not in the {BexConstants.SegmentName.ToLower()} {BexConstants.SublineAllocationName.ToLower()}";
                messages.AppendLine($"{name} appears in a {BexConstants.HazardProfileName.ToLower()} {butNot}");
            }

            if (messages.Length == 0) return;

            stringBuilder.AppendLine($"{segmentName.ConnectWithDash(Name)}");
            stringBuilder.Append(messages);
            stringBuilder.AppendLine();
        }

        public int InterDisplayOrder { get; set; }
        public int IntraDisplayOrder { get; set; }

        public IList<long?> SublineIds { get; set; }
        public int Id { get; set; }
        
        protected static void CheckAmountIsNonNegative(StringBuilder messages, string location, double? amount, string name)
        {
            if (!amount.HasValue || amount.Value >= 0) return;

            var message = $"{name.ToStartOfSentence()} <{amount:N0}> is less than 0 in {location}";
            messages.AppendLine(message);
        }
    }
}