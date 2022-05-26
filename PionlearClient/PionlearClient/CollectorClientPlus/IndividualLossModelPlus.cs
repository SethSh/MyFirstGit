using System;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.Model;

namespace PionlearClient.CollectorClientPlus
{
    public class IndividualLossModelPlus : IndividualLossModel, IModel
    {
        public int RowId { get; set; }
        public int RowNumber { get; set; }
        public string Name { get; set; }
        public long? SourceId { get; set; }
        public long? PredecessorSourceId { get; set; }
        public bool IsDirty { get; set; }
        public Guid Guid { get; set; }
        public long? DataSetId { get; set; }
        public DateTime? SourceTimestamp { get; set; }

        public bool IsDateValid(string historicalPeriodType)
        {
            switch (historicalPeriodType)
            {
                case "1": return AccidentDate.HasValue;
                case "2": return PolicyDate.HasValue;
                case "3": return ReportedDate.HasValue;
                default: throw new ArgumentOutOfRangeException($"Can't find historical period type");
            }
        }

        public bool IsAnyContent()
        {
            return !string.IsNullOrEmpty(OccurrenceId)
                   || !string.IsNullOrEmpty(ClaimNumber)
                   || !string.IsNullOrEmpty(EventCode)
                   || !string.IsNullOrEmpty(LossDescription)
                   || IsAnyPolicyLimitOrAttachment()
                   || IsAnyDate()
                   || IsAnyLossAmount();
        }

        public bool IsEqualTo(IndividualLossModel otherLoss)
        {
            //don't want a mismatch with event code when comparing null and "" - both IsNullOrEmpties should be a match
            //added check for all strings
            if (!IsStringEqualTo(OccurrenceId, otherLoss.OccurrenceId)) return false;
            if (!IsStringEqualTo(ClaimNumber, otherLoss.ClaimNumber)) return false;
            if (!IsStringEqualTo(EventCode, otherLoss.EventCode)) return false;
            if (!IsStringEqualTo(LossDescription, otherLoss.LossDescription)) return false;
           
            if (!PolicyLimitAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PolicyLimitAmount)) return false;
            if (!PolicyAttachmentAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PolicyAttachmentAmount)) return false;

            if (!AccidentDate.IsDateEqualIncludingNull(otherLoss.AccidentDate)) return false;
            if (!PolicyDate.IsDateEqualIncludingNull(otherLoss.PolicyDate)) return false;
            if (!ReportedDate.IsDateEqualIncludingNull(otherLoss.ReportedDate)) return false;

            if (!PaidLossAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PaidLossAmount)) return false;
            if (!PaidAlaeAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PaidAlaeAmount)) return false;
            if (!PaidCombinedAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.PaidCombinedAmount)) return false;

            if (!ReportedLossAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.ReportedLossAmount)) return false;
            if (!ReportedAlaeAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.ReportedAlaeAmount)) return false;
            if (!ReportedCombinedAmount.IsEpsilonEqualIncludingNullAndNaN(otherLoss.ReportedCombinedAmount)) return false;

            return true;
        }

        private static bool IsStringEqualTo(string first, string second)
        {
            if (string.IsNullOrEmpty(first) && string.IsNullOrEmpty(second)) return true;
            return first == second;
        }

        private bool IsAnyPolicyLimitOrAttachment()
        {
            return PolicyLimitAmount.HasValue || PolicyAttachmentAmount.HasValue;
        }

        private bool IsAnyDate()
        {
            return AccidentDate.HasValue
                   || ReportedDate.HasValue
                   || PolicyDate.HasValue;
        }

        private bool IsAnyLossAmount()
        {
            return PaidLossAmount.HasValue
                   || PaidAlaeAmount.HasValue
                   || PaidCombinedAmount.HasValue
                   || ReportedLossAmount.HasValue
                   || ReportedAlaeAmount.HasValue
                   || ReportedCombinedAmount.HasValue;
        }
    }
}
