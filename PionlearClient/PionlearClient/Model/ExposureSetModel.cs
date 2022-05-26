using System.Collections.Generic;
using System.Linq;
using System.Text;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.CollectorClientPlus;

// ReSharper disable PossibleInvalidOperationException

namespace PionlearClient.Model
{
    public class ExposureSetModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.ExposureSetName;
        public IList<ExposureModelPlus> Items { get; set; }
        public short ExposureBaseId { get; set; }

        public override StringBuilder Validate(IList<CollectorApi.Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();

            var exposureMessages = ValidateExposures();
            if (exposureMessages.Length > 0)
            {
                messages.Append(exposureMessages);
            }

            return messages;
        }

        public override StringBuilder PerformQualityControl()
        {
            //do nothing
            return new StringBuilder();
        }

        public ExposureSetModelPlus Map()
        {
            return new ExposureSetModelPlus
            {
                Id = SourceId,
                Name = Name,
                ExposureBaseId = ExposureBaseId,
                Sublines = SublineIds.Select(id => id.Value).ToList()
            };
        }

        private StringBuilder ValidateExposures()
        {
            var messages = new StringBuilder();

            foreach (var item in Items)
            {
                CheckAmountIsNonNegative(messages, item.Location, item.Amount, BexConstants.ExposureSetName);
            }

            return messages;
        }
    }
}
