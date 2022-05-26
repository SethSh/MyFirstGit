using System.Collections.Generic;
using System.Linq;
using System.Text;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.CollectorClientPlus;
using PionlearClient.Extensions;

namespace PionlearClient.Model
{
    public class RateChangeSetModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.RateChangeSetName;
        public CollectorApi.RateChangeSetModel Map()
        {
            return new CollectorApi.RateChangeSetModel
            {
                Id = SourceId,
                Name = Name,
                // ReSharper disable once PossibleInvalidOperationException
                Sublines = SublineIds.Select(id => id.Value).ToList()
            };
        }
        
        public override StringBuilder Validate(IList<CollectorApi.Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();
            
            var rateChangeMessages = ValidateRateChanges();
            if (rateChangeMessages.Length > 0)
            {
                messages.Append(rateChangeMessages);
            }

            return messages;
        }
        
        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            return messages;
        }

        public List<RateChangeModelPlus> Items { get; set; }

        private StringBuilder ValidateRateChanges()
        {
            const int earlierValidYear = 1950;
            const double minimumRateChange = -1;
            var messages = new StringBuilder();
            
            foreach (var item in Items)
            {
                var rowNumberAsString = item.RowNumber.ToString("N0");
                var location = $"row {rowNumberAsString}";
                
                if (item.EffectiveDate.Year < earlierValidYear)
                {
                    messages.AppendLine($"Year of {BexConstants.RateChangeName.ToLower()} date " +
                                        $"can't be < {earlierValidYear} in {location}");
                }

                if (!item.EffectiveDate.IsWithinAcceptableDateRange())
                {
                    var message = $"Date <{item.EffectiveDate:d}> doesn't fall within acceptable year range in {location}";
                    messages.AppendLine(message);
                }


                if (item.Value <= minimumRateChange)
                {
                    messages.AppendLine($"{BexConstants.RateChangeName.ToStartOfSentence()} " +
                                        $"must be > {minimumRateChange:P0} in {location}");
                }

            }

            return messages;
        }

    }
}
