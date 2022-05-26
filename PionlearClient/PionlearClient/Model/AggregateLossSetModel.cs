using System.Collections.Generic;
using System.Linq;
using System.Text;
using PionlearClient.CollectorClientPlus;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;
// ReSharper disable PossibleInvalidOperationException

namespace PionlearClient.Model
{
    public class AggregateLossSetModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.AggregateLossSetName;
        public List<AggregateLossModelPlus> Items { get; set; }
        public bool IsCombinedLossAndAlae { get; set; }
        public bool IsPaidAvailable { get; set; }
        public override StringBuilder Validate(IList<CollectorApi.Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();
            
            var lossValidation = ValidateAggregateLosses();
            if (lossValidation.Length > 0)
            {
                messages.Append(lossValidation);
            }
            
            return messages;
        }

        private StringBuilder ValidateAggregateLosses()
        {
            var messages = new StringBuilder();

            foreach (var item in Items)
            {
                var rowNumberAsString = item.RowNumber.ToString("N0");
                var location = $"row {rowNumberAsString}";
                
                CheckAmountIsNonNegative(messages, location, item.ReportedLossAmount, BexConstants.ReportedLossName);
                CheckAmountIsNonNegative(messages, location, item.ReportedAlaeAmount, BexConstants.ReportedAlaeName);
                CheckAmountIsNonNegative(messages, location, item.ReportedCombinedAmount, BexConstants.ReportedLossAndAlaeName);

                CheckAmountIsNonNegative(messages, location, item.PaidLossAmount, BexConstants.PaidLossName);
                CheckAmountIsNonNegative(messages, location, item.PaidAlaeAmount, BexConstants.PaidAlaeName);
                CheckAmountIsNonNegative(messages, location, item.PaidCombinedAmount, BexConstants.PaidLossAndAlaeName);
            }

            return messages;
        }
        
        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            return messages;
        }

        public AggregateLossSetModelPlus Map()
        {
            return new AggregateLossSetModelPlus
            {
                Id = SourceId,
                Name = Name,
                CombinedLossAndAlae = IsCombinedLossAndAlae,
                Sublines = SublineIds.Select(id => id.Value).ToList()
            };
        }
    }
}