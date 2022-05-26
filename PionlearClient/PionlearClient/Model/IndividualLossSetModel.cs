using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PionlearClient.CollectorClientPlus;
using PionlearClient.Extensions;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;
// ReSharper disable PossibleInvalidOperationException

namespace PionlearClient.Model
{
    public class IndividualLossSetModel : BaseSourceComponentModel
    {
        public int SubjectPolicyAlaeTreatment { get; set; }
        public override string FriendlyName => BexConstants.IndividualLossSetName;
        public bool IsCombinedLossAndAlae { get; set; }
        public int? Threshold { get; set; }

        public List<IndividualLossModelPlus> Items { get; set; }

        private const string WarningOccurrenceId = "0";

        public override StringBuilder Validate(IList<CollectorApi.Allocation> riskClassValidSublines)
        {
            var messages = new StringBuilder();

            var lossMessages = ValidateLosses();
            if (Threshold.HasValue && Threshold < 0) messages.AppendLine($"{BexConstants.ThresholdName.ToStartOfSentence()} " +
                                                                           $"value <{Threshold:N0}> can't be negative");
            if (lossMessages.Length > 0)
            {
                messages.Append(lossMessages);
            }

            return messages;
        }

        public IndividualLossSetModelPlus Map()
        {
            return new IndividualLossSetModelPlus
            {
                Id = SourceId,
                Name = Name,
                Threshold = Threshold,
                CombinedLossAndAlae = IsCombinedLossAndAlae,
                DefaultEvaluationDate = DateTime.Today,
                IncludedPolicyProperties = Items.Any(item => item.PolicyLimitAmount.HasValue || item.PolicyAttachmentAmount.HasValue),
                Sublines = SublineIds.Select(id => id.Value).ToList()
            };
        }
        
        public override StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            if (!Threshold.HasValue) messages.AppendLine($"{BexConstants.ThresholdName.ToStartOfSentence()} is blank");

            foreach (var item in Items)
            {
                var rowNumberAsString = item.RowNumber.ToString("N0");
                var location = $"row {rowNumberAsString}";
                CheckOccurrenceIsNotZero(messages, location, item);

                CheckAmountIsNonNegative(messages, location, item.ReportedLossAmount, BexConstants.ReportedLossName);
                CheckAmountIsNonNegative(messages, location, item.ReportedAlaeAmount, BexConstants.ReportedAlaeName);
                CheckAmountIsNonNegative(messages, location, item.ReportedCombinedAmount, BexConstants.ReportedLossAndAlaeName);

                CheckAmountIsNonNegative(messages, location, item.PaidLossAmount, BexConstants.PaidLossName);
                CheckAmountIsNonNegative(messages, location, item.PaidAlaeAmount, BexConstants.PaidAlaeName);
                CheckAmountIsNonNegative(messages, location, item.PaidCombinedAmount, BexConstants.PaidLossAndAlaeName);
                
                CheckLimitGreaterThanReported(messages, location, item);
                CheckReportedLossGreaterThanPaidLoss(messages, location, item);
                CheckReportedAlaeGreaterThanPaidAlae(messages, location, item);
                CheckReportedLossAndAlaeGreaterThanPaidLossAndAlae(messages, location, item);
            }

            return messages;
        }

        public IEnumerable<IndividualLossModelPlus> SlotLosses(int historicalPeriodType, DateTime startDate, DateTime endDate)
        {
            switch (historicalPeriodType)
            {
                case 1:
                    return Items.Where(item => item.AccidentDate.Value.ConvertFromDateTimeOffset() >= startDate
                                               && item.AccidentDate.Value.ConvertFromDateTimeOffset() <= endDate);
                case 2:
                    return Items.Where(item => item.PolicyDate.Value.ConvertFromDateTimeOffset() >= startDate
                                               && item.PolicyDate.Value.ConvertFromDateTimeOffset() <= endDate);
                case 3:
                    return Items.Where(item => item.ReportedDate.Value.ConvertFromDateTimeOffset() >= startDate &&
                                               item.ReportedDate.Value.ConvertFromDateTimeOffset() <= endDate);
            }

            throw new ArgumentOutOfRangeException($"Can't recognize year type");
        }
        
        private StringBuilder ValidateLosses()
        {
            var messages = new StringBuilder();
            CheckAnyReportedLoss(messages);
            CheckAnyClaimDuplicates(messages);
                        
            foreach (var item in Items)
            {
                var rowNumberAsString = item.RowNumber.ToString("N0");
                var location = $"row {rowNumberAsString}";
                CheckClaimIdHasValue(messages, location, item);
                CheckAmountIsNonNegative(messages, location, item.PolicyLimitAmount, BexConstants.PolicyLimitName);
                CheckAmountIsNonNegative(messages, location, item.PolicyAttachmentAmount, BexConstants.PolicySirName);
                CheckDataIsInAcceptableRangeAmount(messages, location, item);
            }

            return messages;
        }

        private void CheckAnyClaimDuplicates(StringBuilder messages)
        {
            const int displayMaximumLength = 10;
            const int displayMaximumRowLength = 10;
            var claims = Items.Where(claim => !string.IsNullOrEmpty(claim.ClaimNumber));

            var duplicateClaimIds = claims.Select(claim=>claim.ClaimNumber).Duplicates().ToList();
            if (!duplicateClaimIds.Any()) return;

            foreach (var duplicateClaimId in duplicateClaimIds.Take(displayMaximumLength))
            {
                var matchingClaims = Items.Where(item => item.ClaimNumber == duplicateClaimId);
                var rowNumbers = matchingClaims.Select(item => item.RowNumber).ToList();
                var rowCount = rowNumbers.Count;

                var claimAsString  = $"<{duplicateClaimId}> in rows: ";
                if (rowCount > displayMaximumRowLength)
                {
                    claimAsString += string.Join(", ", rowNumbers.Take(displayMaximumRowLength).ToArray()) + " ...";
                }
                else
                {
                    claimAsString += string.Join(", ", rowNumbers.ToArray());
                }
                
                messages.AppendLine($"Contains duplicate {BexConstants.ClaimIdName} {claimAsString}");
            }

            if (duplicateClaimIds.Count <= displayMaximumLength) return;

            var notDisplayedClaimCount = duplicateClaimIds.Count - displayMaximumLength;
            messages.AppendLine(
                notDisplayedClaimCount == 1
                    ? $"There is {notDisplayedClaimCount:N0} additional duplicate not displayed"
                    : $"There are {notDisplayedClaimCount:N0} additional duplicates not displayed");
        }

        private void CheckAnyReportedLoss(StringBuilder messages)
        {
            if (Items.Count == 0) return;
            if (IsCombinedLossAndAlae)
            {
                if (Items.All(item => !item.ReportedCombinedAmount.HasValue))
                {
                    messages.AppendLine($"{BexConstants.ReportedLossAndAlaeName} are all blank");
                }
            }
            else
            {
                if (Items.All(item => !item.ReportedLossAmount.HasValue && !item.ReportedAlaeAmount.HasValue))
                {
                    messages.AppendLine($"{BexConstants.ReportedLossName} and {BexConstants.ReportedAlaeName} are all blank");
                }
            }
        }

        private static void CheckReportedLossGreaterThanPaidLoss(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            if (!item.PaidLossAmount.IsGreaterThan(item.ReportedLossAmount)) return;

            var message = $"{BexConstants.PaidLossName.ToStartOfSentence()} <{item.PaidLossAmount:N0}> " +
                          $"is greater than {BexConstants.ReportedLossName} <{item.ReportedLossAmount:N0}> in {location}";
            messages.AppendLine(message);
        }

        private static void CheckReportedAlaeGreaterThanPaidAlae(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            if (!item.PaidAlaeAmount.IsGreaterThan(item.ReportedAlaeAmount)) return;

            var message = $"{BexConstants.PaidAlaeName.ToStartOfSentence()} <{item.PaidAlaeAmount:N0}> " +
                          $"is greater than {BexConstants.ReportedAlaeName.ToLower()} <{item.ReportedAlaeAmount:N0}> in {location}";
            messages.AppendLine(message);
        }

        private static void CheckReportedLossAndAlaeGreaterThanPaidLossAndAlae(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            if (!item.PaidCombinedAmount.IsGreaterThan(item.ReportedCombinedAmount)) return;

            var message = $"{BexConstants.PaidLossAndAlaeName.ToStartOfSentence()} <{item.PaidCombinedAmount:N0}> " +
                          $"is greater than {BexConstants.ReportedLossAndAlaeName.ToLower()} <{item.ReportedCombinedAmount:N0}> " +
                          $"in {location}";
            messages.AppendLine(message);
        }

        private void CheckLimitGreaterThanReported(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            if (SubjectPolicyAlaeTreatment == 1)
            {
                CheckReportedLossPlusAlaeNotGreaterThanLimit(messages, location, item);
            }
            else
            {
                CheckReportedLossNotGreaterThanLimit(messages, location, item);
            }
        }

        private static void CheckClaimIdHasValue(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            var hasClaimId = !string.IsNullOrEmpty(item.ClaimNumber);
            if (hasClaimId) return;

            var message = $"{BexConstants.ClaimIdName} is blank in {location}";
            messages.AppendLine(message);
        }

        private static void CheckOccurrenceIsNotZero(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            var id = item.OccurrenceId;
            if (id == null || !id.Equals(WarningOccurrenceId)) return;

            var message = $"{BexConstants.OccurrenceIdName} is <{WarningOccurrenceId}> in {location}" +
                          $" (This could an undesirable result of a linking to a blank call).";
            messages.AppendLine(message);
        }

        private static void CheckReportedLossNotGreaterThanLimit(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            if (!item.ReportedLossAmount.IsGreaterThan(item.PolicyLimitAmount)) return;

            var message = $"{BexConstants.ReportedLossName.ToStartOfSentence()} <{item.ReportedLossAmount:N0}> " +
                          $"is greater than {BexConstants.PolicyLimitName.ToLower()} <{item.PolicyLimitAmount:N0}> in {location}";
            messages.AppendLine(message);
        }

        private static void CheckReportedLossPlusAlaeNotGreaterThanLimit(StringBuilder messages, string location, IndividualLossModelPlus item)
        {
            if (!item.ReportedLossAmount.HasValue && !item.ReportedAlaeAmount.HasValue) return;

            var reportLossPlusAlae = item.ReportedLossAmount ?? 0 + item.ReportedAlaeAmount ?? 0;
            if (!reportLossPlusAlae.IsGreaterThan(item.PolicyLimitAmount)) return;
            
            var message = $"{BexConstants.ReportedLossAndAlaeName} <{reportLossPlusAlae:N0}> " +
                          $"is greater than {BexConstants.PolicyLimitName.ToLower()} <{item.PolicyLimitAmount:N0}> in {location}";
            messages.AppendLine(message);
        }
        
        private static void CheckDataIsInAcceptableRangeAmount(StringBuilder messages, string location, IndividualLossModelPlus loss)
        {
            var messageSuffix = $"doesn't fall within acceptable year range in {location}"; 
            
            if (!loss.ReportedDate.IsWithinAcceptableDateRange())
            {
                var message = loss.ReportedDate.HasValue
                    ? $"{BexConstants.ReportDateName.ToStartOfSentence()} <{loss.ReportedDate:d}> {messageSuffix}"
                    : $"{BexConstants.ReportDateName.ToStartOfSentence()} {messageSuffix}";
                messages.AppendLine(message);
            }

            if (!loss.AccidentDate.IsWithinAcceptableDateRange())
            {
                var message = loss.AccidentDate.HasValue
                    ? $"{BexConstants.AccidentDateName.ToStartOfSentence()} <{loss.AccidentDate:d}> {messageSuffix}"
                    : $"{BexConstants.AccidentDateName.ToStartOfSentence()} {messageSuffix}";
                messages.AppendLine(message);
            }

            if (!loss.PolicyDate.IsWithinAcceptableDateRange())
            {
                var message = loss.PolicyDate.HasValue
                    ? $"{BexConstants.PolicyDateName.ToStartOfSentence()} <{loss.PolicyDate:d}> {messageSuffix}"
                    : $"{BexConstants.PolicyDateName.ToStartOfSentence()} {messageSuffix}";
                messages.AppendLine(message);
            }
        }
    }
}
