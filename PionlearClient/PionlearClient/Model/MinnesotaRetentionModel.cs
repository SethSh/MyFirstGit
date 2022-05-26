using System;
using System.Collections.Generic;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.Model
{
    public class MinnesotaRetentionModel : BaseSourceComponentModel
    {
        public override string FriendlyName => BexConstants.WorkersCompClassCodeProfileName;
        public long? RetentionId { get; set; }
        public double RetentionValue { get; set; }

        
        public override StringBuilder Validate(IList<Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();
            
            if (RetentionValue < 0)
            {
                messages.AppendLine($"Change Minnesota Retention <{RetentionValue:N2}> to a non-negative number");
            }
            return messages;
        }

        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();

            return messages;
        }

        public MunichRe.Bex.ApiClient.ClientApi.MinnesotaRetentionModel Map()
        {
            return new MunichRe.Bex.ApiClient.ClientApi.MinnesotaRetentionModel
            {
                Id = RetentionId.Value,
                RetentionAmount = Convert.ToInt64(RetentionValue)
            };
        }
    }
}