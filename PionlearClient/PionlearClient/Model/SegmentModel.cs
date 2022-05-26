using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.Model
{
    public class SegmentModel : BaseSourceModel
    {
        public SegmentModel()
        {
            Guid = Guid.NewGuid();
            IsDirty = true;
            IsUmbrella = false;
            Models = new List<ISourceComponentModel>();
        }
        
        public double ProspectiveExposureAmount { get; set; }
        public int ProspectiveExposureBasisCode { get; set; }
        public int HistoricalExposureBasisCode { get; set; }
        public int HistoricalPeriodType { get; set; }
        public IList<Allocation> ValidSublines { get; set; }
        public IList<SubmissionSegmentPeriod> ValidPeriods { get; set; }
        
        public IList<Allocation> UmbrellaTypeAllocations { get; set; }


        public IList<ISourceComponentModel> Models { get; set; }
        public IList<PolicyModel> PolicyModels => Models.OfType<PolicyModel>().ToList();
        public IList<StateModel> StateModels => Models.OfType<StateModel>().ToList();
        public IList<HazardModel> HazardModels => Models.OfType<HazardModel>().ToList();
        
        public IList<ConstructionTypeModel> ConstructionTypeModels => Models.OfType<ConstructionTypeModel>().ToList();
        public IList<OccupancyTypeModel> OccupancyTypeModels => Models.OfType<OccupancyTypeModel>().ToList();
        public IList<ProtectionClassModel> ProtectionClassModels => Models.OfType<ProtectionClassModel>().ToList();
        public IList<TotalInsuredValueModel> TotalInsuredValueModels => Models.OfType<TotalInsuredValueModel>().ToList();

        
        public IList<WorkersCompStateAttachmentModel> WorkersCompStateAttachmentModels => Models.OfType<WorkersCompStateAttachmentModel>().ToList();
        public IList<WorkersCompStateHazardGroupModel> WorkersCompStateHazardGroupModels => Models.OfType<WorkersCompStateHazardGroupModel>().ToList();
        public IList<WorkersCompClassCodeModel> WorkersCompClassCodeModels => Models.OfType<WorkersCompClassCodeModel>().ToList();


        public IList<ExposureSetModel> ExposureSetModels => Models.OfType<ExposureSetModel>().ToList();
        public IList<AggregateLossSetModel> AggregateLossSetModels => Models.OfType<AggregateLossSetModel>().ToList();
        public IList<IndividualLossSetModel> IndividualLossSetModels => Models.OfType<IndividualLossSetModel>().ToList();
        public IList<RateChangeSetModel> RateChangeSetModels => Models.OfType<RateChangeSetModel>().ToList();

        
        public bool IsUmbrella { get; set; }
        public int SubjectPolicyAlaeTreatment { get; set; }
        public bool IsSublineAllocationPercent { get; set; }
        public bool IncludePaidInAggregateLoss { get; set; }
        public bool IncludePaidInIndividualLoss { get; set; }
        public long? MinnesotaRetentionId { get; set; }
        public bool IsProperty { get; set; }
        
        public SubmissionSegmentModel Map()
        {
            var submissionSegmentModel = new SubmissionSegmentModel();
            MapHeader(submissionSegmentModel);
            MapUmbrellaTypeAllocations(submissionSegmentModel);
            MapValidSublines(submissionSegmentModel);
            MapValidPeriods(submissionSegmentModel);
            return submissionSegmentModel;
        }
        
        public void MapHeader(SubmissionSegmentModel submissionSegmentModel)
        {
            submissionSegmentModel.Id = SourceId;
            submissionSegmentModel.IsUmbrella = IsUmbrella;
            submissionSegmentModel.IsAllocationPercent = IsSublineAllocationPercent;
            submissionSegmentModel.Name = Name;
            submissionSegmentModel.PeriodType = HistoricalPeriodType;
            submissionSegmentModel.SubjectBaseAmount = ProspectiveExposureAmount;
            submissionSegmentModel.SubjectBaseUnitOfMeasure = ProspectiveExposureBasisCode;
            submissionSegmentModel.SubjectPolicyAlaeTreatment = SubjectPolicyAlaeTreatment;
            submissionSegmentModel.HistoricalSubjectBaseUnitOfMeasure = HistoricalExposureBasisCode;
            submissionSegmentModel.IncludePaidInAggregateLoss = IncludePaidInAggregateLoss;
            submissionSegmentModel.IncludePaidInIndividualLoss = IncludePaidInIndividualLoss;
            submissionSegmentModel.MinnesotaRetentionId = MinnesotaRetentionId;
        }

        public void MapUmbrellaTypeAllocations(SubmissionSegmentModel submissionSegmentModel)
        {
            if (UmbrellaTypeAllocations == null)
            {
                submissionSegmentModel.UmbrellaTypeAllocations = new List<Allocation>();
                return;
            }

            var allocations = new List<Allocation>();
            UmbrellaTypeAllocations.ForEach(u =>
            {
                allocations.Add(new Allocation {Id = u.Id, Value = u.Value});
            });
            submissionSegmentModel.UmbrellaTypeAllocations = allocations;
        }

        public void MapValidSublines(SubmissionSegmentModel submissionSegmentModel)
        {
            var allocations = new List<Allocation>();
            ValidSublines.ForEach(subline =>
            {
                allocations.Add(new Allocation { Id = subline.Id, Value = subline.Value });
            });
            submissionSegmentModel.ValidSublines = allocations;
        }

        public void MapValidPeriods(SubmissionSegmentModel submissionSegmentModel)
        {
            var periods = new List<SubmissionSegmentPeriod>();
            ValidPeriods.ForEach(period =>
            {
                periods.Add(new SubmissionSegmentPeriod { StartDate = period.StartDate, EndDate = period.EndDate, EvaluationDate = period.EvaluationDate});
            });
            submissionSegmentModel.ValidPeriods = periods;
        }

        public StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            if (IsUmbrella)
            {
                CheckUmbrellaAllocationConsistency(messages);
            }
            CheckPeriodsHaveNoGaps(messages);
            CheckEvaluationDatesArePriorToNow(messages);
            CheckAggregateLossGreaterThanOrEqualToSumOfIndividualLoss(messages);
            CheckSubjectPolicyAppliesTo(messages);
            return messages;
        }

        private void CheckUmbrellaAllocationConsistency(StringBuilder messages)
        {
            var separatedUmbrellaTypeIds = PolicyModels
                .Where(model => model.UmbrellaTypeId.HasValue && (model.UmbrellaTypeId ?? 0) != UmbrellaTypesFromBex.GetPersonalCode())
                .Select(model => model.UmbrellaTypeId.Value);

            separatedUmbrellaTypeIds.ForEach(id =>
            {
                var umbrellaAllocation = UmbrellaTypeAllocations.Single(alloc => alloc.Id == id);
                if (!umbrellaAllocation.Value.IsEqual(0)) return;

                var name = UmbrellaTypesFromBex.ReferenceData.Single(umb => umb.UmbrellaTypeCode == id).UmbrellaTypeName;
                messages.AppendLine($"The {name} has a {0d:p0} {BexConstants.UmbrellaAllocationName.ToLower()} allocation.");
            });
        }

        private void ValidateSublineAllocations(StringBuilder messages)
        {
            //for now this bool will always to true
            if (!IsSublineAllocationPercent) return;

            var allocations = ValidSublines.Select(sub => sub.Value);
            foreach (var allocation in allocations)
            {
                if (double.IsNaN(allocation))
                {
                    messages.AppendLine($"Change {BexConstants.SublineAllocationName.ToLower()} to a number");
                }
                else
                {
                    if (allocation < 0 || allocation > 1) messages.AppendLine($"Change {BexConstants.SublineAllocationName.ToLower()} value <{allocation:P2}> to a number between 0 and 1");
                }
            }
        }

        private void ValidateSublineAllocationSum(StringBuilder messages)
        {
            //for now this book will always to true
            if (!IsSublineAllocationPercent) return;

            var tolerance = NumericalConstants.ValidationProfileTolerance;
            var valueSum = ValidSublines.Sum(sub => sub.Value);
            if (!valueSum.IsEpsilonEqual(1d, tolerance))
            {
                messages.AppendLine($"The {BexConstants.SublineAllocationName.ToLower()} percent sum <{valueSum:P4}> isn't within {tolerance:P4} of {1:P4}");
            }
        }

        private void ValidateUmbrellaSum(StringBuilder messages)
        {
            if (!IsUmbrella) return;

            var tolerance = NumericalConstants.ValidationProfileTolerance;
            var personalCode = UmbrellaTypesFromBex.GetPersonalCode();
            
            var commercialUmbrellaTypes = UmbrellaTypeAllocations.Where(alloc => alloc.Id != personalCode).ToList();
            if (!commercialUmbrellaTypes.Any()) return;

            var valueSum = commercialUmbrellaTypes.Sum(item => item.Value);
            if (!valueSum.IsEpsilonEqual(1, tolerance))
            {
                messages.AppendLine(
                    $"The {BexConstants.UmbrellaAllocationName.ToLower()} percent sum <{valueSum:P4}> isn't within {tolerance:P4} of {1:P4}");
            }
        }

        private void ValidateUmbrellaAllocations(StringBuilder messages)
        {
            if (!IsUmbrella) return;

            var allocations = UmbrellaTypeAllocations.Select(sub => sub.Value);
            foreach (var allocation in allocations)
            {
                if (double.IsNaN(allocation))
                {
                    messages.AppendLine($"Change {BexConstants.UmbrellaAllocationName.ToLower()} to a number");
                }
                else
                {
                    if (allocation < 0 || allocation > 1) messages.AppendLine($"Change {BexConstants.UmbrellaAllocationName.ToLower()} value <{allocation:P2}> to a number between 0 and 1");
                }
            }
        }

        private void ValidateAtLeastOneOfIndividualOrAggregateLosses(StringBuilder messages)
        {
            if (!ExposureSetModels.Any(x => x.Items.Count > 0))
            {
                if (IndividualLossSetModels.Any(x => x.Items.Count > 0)) {
                    messages.AppendLine($"{BexConstants.SegmentName.ToStartOfSentence()} with no historical data cannot have any {BexConstants.IndividualLossSetName.ToLower()}");
                }
                else return;
            }

            if (IndividualLossSetModels.Any(x => x.Items.Count > 0) || AggregateLossSetModels.Any(x => x.Items.Count > 0)) return;

            messages.AppendLine($"{BexConstants.SegmentName.ToStartOfSentence()} must have at least one of {BexConstants.AggregateLossSetName.ToLower()} or {BexConstants.IndividualLossSetName.ToLower()}");
        }

        private void CheckAggregateLossGreaterThanOrEqualToSumOfIndividualLoss(StringBuilder messages)
        {
            if (!IndividualLossSetModels.Any()) return;
            if (!IndividualLossSetModels.Any(model => model.Items.Any())) return;

            if (!AggregateLossSetModels.Any()) return;
            if (!AggregateLossSetModels.Any(model => model.Items.Any())) return;
            
            var isAggregateLossCombined = AggregateLossSetModels.First().Items.Any(lossSet => lossSet.ReportedCombinedAmount != null);
            var isIndividualLossCombined = IndividualLossSetModels.First().Items.Any(lossSet => lossSet.ReportedCombinedAmount != null);

            var miniMessages = new StringBuilder();
            foreach (var period in ValidPeriods.OrderBy(p => p.StartDate))
            {
                var periodAggregateLosses = AggregateLossSetModels
                    .Select(aggModel => aggModel.Items.SingleOrDefault(model => model.StartDate == period.StartDate))
                    .Where(item => item != null);

                var aggregateReportedAmount = isAggregateLossCombined
                    ? periodAggregateLosses.Sum(amount => amount.ReportedCombinedAmount.GetValueOrDefault())
                    : periodAggregateLosses.Sum(amount => amount.ReportedLossAmount.GetValueOrDefault()
                                                        + amount.ReportedAlaeAmount.GetValueOrDefault());
                
                var individualReportedAmount = CalculateSumOfIndividualLoss(period.StartDate.Date, period.EndDate.Date, isIndividualLossCombined);

                //using decimals is tricky because of the way floats get added
                //so instead just compare the left of the decimal as this is only for qc
                var individualReportedAmountAsLong = individualReportedAmount.RoundToLong();
                var aggregateReportedAmountAsLong = aggregateReportedAmount.RoundToLong();
                if (!(individualReportedAmountAsLong > aggregateReportedAmountAsLong)) continue;

                var dateMessage = period.StartDate.ToString("d").ConnectWithDash(period.EndDate.ToString("d"));
                var valueMessage = $"{BexConstants.AggregateLossSetName.ToLower()} {aggregateReportedAmount:N0} is less than " +
                                   $"{BexConstants.IndividualLossSetName.ToLower()} {individualReportedAmount:N0} ";
                miniMessages.AppendLine($"{dateMessage}: {valueMessage}");
            }

            if (miniMessages.Length <= 0) return;

            if (messages.Length > 0) messages.AppendLine();
            messages.AppendLine($"{BexConstants.AggregateLossSetName} / {BexConstants.IndividualLossSetName} Inconsistency");
            messages.Append(miniMessages);
        }

        private void CheckSubjectPolicyAppliesTo(StringBuilder messages)
        {
            if (IsProperty && SubjectPolicyAlaeTreatment != BexConstants.LossOnlySubjectPolicyAlaeTreatmentId)
            {
                messages.Append("Property segment has the Subject Policy Limits Apply To set to Loss & ALAE Combined");
            }
        }

        private double CalculateSumOfIndividualLoss(DateTime startDate, DateTime endDate, bool isLossCombined)
        {
            var reportedAmount = 0d;
            foreach (var individualLossSetModel in IndividualLossSetModels)
            {
                var periodLosses = individualLossSetModel.SlotLosses(HistoricalPeriodType, startDate, endDate);
                var periodReportedAmount = isLossCombined
                    ? periodLosses.Sum(item => item.ReportedCombinedAmount.GetValueOrDefault())
                    : periodLosses.Sum(item => item.ReportedLossAmount.GetValueOrDefault()
                                               + item.ReportedAlaeAmount.GetValueOrDefault());
                reportedAmount += periodReportedAmount;
            }
            return reportedAmount;
        }

        private void CheckPeriodsHaveNoGaps(StringBuilder messages)
        {
            var miniMessages = new StringBuilder();
            DateTime? previousPeriodStart = null;
            DateTime? previousPeriodEnd = null;
            foreach (var period in ValidPeriods.OrderBy(p => p.StartDate))
            {
                if (previousPeriodEnd.HasValue)
                {
                    if ((period.StartDate - previousPeriodEnd.Value).TotalDays > 1)
                    {
                        var previous = $"{previousPeriodStart:d}-{previousPeriodEnd:d}";
                        var current = $"{period.StartDate:d}-{period.EndDate:d}";
                        miniMessages.AppendLine($"There's a gap between {previous} and {current}");
                    }
                }
                previousPeriodStart = period.StartDate.Date;
                previousPeriodEnd = period.EndDate.Date;
            }

            if (miniMessages.Length <= 0) return;

            if (messages.Length > 0) messages.AppendLine();
            messages.AppendLine(BexConstants.PeriodName);
            messages.Append(miniMessages);
        }

        private void CheckEvaluationDatesArePriorToNow(StringBuilder messages)
        {
            var miniMessages = new StringBuilder();
            var today = DateTime.Now.Date;
            var evaluateDates = ValidPeriods.OrderBy(p => p.EvaluationDate)
                .Where(p => p.EvaluationDate.Date > today)
                .Select(p =>p.EvaluationDate.Date)
                .Distinct();

            foreach (var evaluationDate in evaluateDates)
            {
                miniMessages.AppendLine($"{evaluationDate:d} is in the future");
            }

            if (miniMessages.Length <= 0) return;

            if (messages.Length > 0) messages.AppendLine();
            messages.AppendLine(BexConstants.EvaluationDateName);
            messages.Append(miniMessages);
        }
        
        public StringBuilder Validate()
        {
            var validation = new StringBuilder();

            ValidateName(validation);
            ValidateProspectiveExposureAmount(validation);
            ValidateSublineAllocations(validation);
            ValidateSublineAllocationSum(validation);
            ValidateCommercialSublinesHaveUmbrellaTypes(validation);
            ValidateSeparatedUmbrellaTypesHaveAllocations(validation);
            ValidateUmbrellaAllocations(validation);
            ValidateUmbrellaSum(validation);
            ValidateChildSublineComposition(validation);
            ValidateChildNotEmpty(validation);
            ValidateUniqueComponentNames(validation);
            ValidatePeriodsDoNotOverlap(validation);
            ValidatePeriodEndsDoNotPrecedeStarts(validation);
            ValidatePeriodStartsPrecedeEvaluations(validation);
            ValidatePeriodsContainAcceptableDates(validation);
            ValidateAtLeastOneOfIndividualOrAggregateLosses(validation);

            return validation;
        }

        private void ValidateSeparatedUmbrellaTypesHaveAllocations(StringBuilder validation)
        {
            var allCommercialSublobIds = SublineCodesFromBex.ReferenceData.Where(x => !x.IsPersonal).Select(x => x.SublineId);
            var commercialAllocations = ValidSublines.Where(sublob => allCommercialSublobIds.Any(sublobId => sublobId == sublob.Id)).ToList();

            if (!IsUmbrella) return;
            if (!commercialAllocations.Any()) return;

            var separatedUmbrellaTypeIds = PolicyModels.Where(model => model.UmbrellaTypeId.HasValue && model.UmbrellaTypeId.Value != UmbrellaTypesFromBex.GetPersonalCode())
                .Select(model => model.UmbrellaTypeId.Value);

            separatedUmbrellaTypeIds.ForEach(id =>
            {
                var umbrellaAllocation = UmbrellaTypeAllocations.SingleOrDefault(alloc => alloc.Id == id);
                if (umbrellaAllocation != null) return;

                var name = UmbrellaTypesFromBex.ReferenceData.Single(umb => umb.UmbrellaTypeCode == id).UmbrellaTypeName;
                validation.AppendLine($"The {name} has no {BexConstants.UmbrellaAllocationName.ToLower()} allocation.");
            });
        }

        private void ValidateName(StringBuilder validation)
        {
            if (string.IsNullOrEmpty(Name)) validation.AppendLine("Name can't be blank");
        }

        private void ValidateProspectiveExposureAmount(StringBuilder validation)
        {
            if (double.IsNaN(ProspectiveExposureAmount)) validation.AppendLine("Prospective exposure amount can't be a non number");
            if (ProspectiveExposureAmount < 0) validation.AppendLine("Prospective exposure amount can't be less than 0");
        }

        private void ValidateCommercialSublinesHaveUmbrellaTypes(StringBuilder validation)
        {
            var allCommercialSublobIds = SublineCodesFromBex.ReferenceData.Where(x => !x.IsPersonal).Select(x => x.SublineId);
            var commercialAllocations = ValidSublines.Where(x => allCommercialSublobIds.Any(y => y == x.Id)).ToList();

            if (!IsUmbrella) return;
            if (!commercialAllocations.Any()) return;

            foreach (var commercialAllocation in commercialAllocations)
            {
                var distributions = PolicyModels.Where(x => x.SublineIds.Contains(commercialAllocation.Id)).ToList();
                if (!distributions.Any()) continue; // orphaned policy sublines is dealt with below

                if (distributions.Count == 1 && !distributions.First().UmbrellaTypeId.HasValue) continue;

                var sublineUmbrellaTypeIds = distributions.Select(x => x.UmbrellaTypeId.GetValueOrDefault());
                var orphanUmbrellaTypeIds = UmbrellaTypeAllocations.Select(x => Convert.ToInt32(x.Id))
                    .Except(sublineUmbrellaTypeIds).ToList();
                if (!orphanUmbrellaTypeIds.Any()) continue;

                var subline = SublineCodesFromBex.ReferenceData.First(x => x.SublineId == commercialAllocation.Id);
                var sublineName = subline.LineOfBusiness.Name.ConnectWithDash(subline.SublineName);
                var orphanUmbrellaTypeNames = orphanUmbrellaTypeIds.Select(x => UmbrellaTypesFromBex.ReferenceData
                    .First(y => y.UmbrellaTypeCode == x).UmbrellaTypeName);
                validation.AppendLine(
                    $"{sublineName} doesn't contain the following umbrella types: {string.Join(", ", orphanUmbrellaTypeNames)} ");
            }

            if (UmbrellaTypeAllocations == null || !UmbrellaTypeAllocations.Any())
            {
                validation.AppendLine("Umbrella profile can't be blank");
            }
        }

        private void ValidateChildSublineComposition(StringBuilder validation)
        {
            // every segment subline appears in exactly one exhibit
            foreach (var sublineId in ValidSublines.Select(x => x.Id))
            {
                var subline = SublineCodesFromBex.ReferenceData.First(x => x.SublineId == sublineId);
                var name = subline.LineOfBusiness.Name.ConnectWithDash(subline.SublineName);

                HazardModels.ValidateSublineComposition(validation, sublineId, name);
                PolicyModels.ValidateSublineComposition(validation, sublineId, name);
                StateModels.ValidateSublineComposition(validation, sublineId, name);
                
                ProtectionClassModels.ValidateSublineComposition(validation, sublineId, name);
                ConstructionTypeModels.ValidateSublineComposition(validation, sublineId, name);
                OccupancyTypeModels.ValidateSublineComposition(validation, sublineId, name);
                TotalInsuredValueModels.ValidateSublineComposition(validation, sublineId, name);

                
                WorkersCompStateAttachmentModels.ValidateSublineComposition(validation, sublineId, name);
                WorkersCompStateHazardGroupModels.ValidateSublineComposition(validation, sublineId, name);
                WorkersCompClassCodeModels.ValidateSublineComposition(validation, sublineId, name);


                ValidateExposureSetModelsSublineComposition(validation, sublineId, name);
                ValidateAggregateLossSetModelsSublineComposition(validation, sublineId, name);
                ValidateIndividualLossSetModelsSublineComposition(validation, sublineId, name);
                ValidateRateChangeSetModelsSublineComposition(validation, sublineId, name);
            }
        }

        private void ValidateRateChangeSetModelsSublineComposition(StringBuilder validation, long sublineId, string name)
        {
            //retrofitting old workbooks that may not have any rate change sets
            if (!RateChangeSetModels.Any()) return;

            var setName = BexConstants.RateChangeSetName.ToLower();
            var lossSets = RateChangeSetModels.Where(model => model.SublineIds.Contains(sublineId)).ToList();
            if (lossSets.Any())
            {
                if (lossSets.Count != 1)
                    validation.AppendLine($"{name} is contained in more than one {setName}");
            }
            else
            {
                validation.AppendLine($"{name} doesn't appear in a {setName}");
            }
        }

        private void ValidateIndividualLossSetModelsSublineComposition(StringBuilder validation, long sublineId, string name)
        {
            var lossSets = IndividualLossSetModels.Where(x => x.SublineIds.Contains(sublineId)).ToList();
            if (lossSets.Any())
            {
                if (lossSets.Count != 1)
                    validation.AppendLine($"{name} is contained in more than one {BexConstants.IndividualLossSetName.ToLower()}");
            }
            else
            {
                validation.AppendLine($"{name} doesn't appear in a {BexConstants.IndividualLossSetName.ToLower()}");
            }
        }

        private void ValidateAggregateLossSetModelsSublineComposition(StringBuilder validation, long sublineId, string name)
        {
            var lossSets = AggregateLossSetModels.Where(x => x.SublineIds.Contains(sublineId)).ToList();
            if (lossSets.Any())
            {
                if (lossSets.Count != 1)
                    validation.AppendLine($"{name} is contained in more than one {BexConstants.AggregateLossSetName.ToLower()}");
            }
            else
            {
                validation.AppendLine($"{name} doesn't appear in a {BexConstants.AggregateLossSetName.ToLower()}");
            }
        }

        private void ValidateExposureSetModelsSublineComposition(StringBuilder validation, long sublineId, string name)
        {
            var exposureSet = ExposureSetModels.Where(x => x.SublineIds.Contains(sublineId)).ToList();
            if (exposureSet.Any())
            {
                if (exposureSet.Count != 1)
                    validation.AppendLine($"{name} is contained in more than one {BexConstants.ExposureSetName.ToLower()}");
            }
            else
            {
                validation.AppendLine($"{name} doesn't appear in a {BexConstants.ExposureSetName.ToLower()}");
            }
        }
        
        private void ValidateChildNotEmpty(StringBuilder validation)
        {
            HazardModels.ValidateEmptiness(validation);
            PolicyModels.ValidateEmptiness(validation);
            StateModels.ValidateEmptiness(validation);

            ProtectionClassModels.ValidateEmptiness(validation);
            ConstructionTypeModels.ValidateEmptiness(validation);
            OccupancyTypeModels.ValidateEmptiness(validation);
            TotalInsuredValueModels.ValidateEmptiness(validation);
            
            //for each each wc profile type: at most one profile, so we don't need to check for some profiles having data and not others
            
            ValidateExposureSetModelsEmptiness(validation);
            ValidateAggregateLossSetModelsEmptiness(validation);
            ValidateIndividualLossSetModelsEmptiness(validation);
        }
        
        private void ValidateExposureSetModelsEmptiness(StringBuilder validation)
        {
            var notEmptyCount = ExposureSetModels.Count(model => !model.Items.Any());
            var emptyCount = ExposureSetModels.Count(model => model.Items.Any());
            if (emptyCount > 0 && notEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.ExposureSetName.ToLower()} contains data");
            }
        }

        private void ValidateAggregateLossSetModelsEmptiness(StringBuilder validation)
        {
            var notEmptyCount = AggregateLossSetModels.Count(model => !model.Items.Any());
            var emptyCount = AggregateLossSetModels.Count(model => model.Items.Any());
            if (emptyCount > 0 && notEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.AggregateLossSetName.ToLower()} contains data");
            }
        }

        private void ValidateIndividualLossSetModelsEmptiness(StringBuilder validation)
        {
            var notEmptyCount = IndividualLossSetModels.Count(model => !model.Items.Any());
            var emptyCount = IndividualLossSetModels.Count(model => model.Items.Any());
            if (emptyCount > 0 && notEmptyCount > 0)
            {
                validation.AppendLine($"Not every {BexConstants.IndividualLossSetName.ToLower()} contains data");
            }
        }

        private void ValidateUniqueComponentNames(StringBuilder validation)
        {
            PolicyModels.Select(model => model.Name).Duplicates()
                .ForEach(name => validation.AppendLine($"The name {name} appears in more than one {BexConstants.PolicyProfileName.ToLower()}"));

            StateModels.Select(model => model.Name).Duplicates()
                .ForEach(name => validation.AppendLine($"The name {name} appears in more than one {BexConstants.StateProfileName.ToLower()}"));

            HazardModels.Select(model => model.Name).Duplicates().
                ForEach(name => validation.AppendLine($"The name {name} appears in more than one {BexConstants.HazardProfileName.ToLower()}"));

            ExposureSetModels.Select(model => model.Name).Duplicates()
                .ForEach(name => validation.AppendLine($"The name {name} appears in more than one {BexConstants.ExposureSetName.ToLower()}"));

            AggregateLossSetModels.Select(model => model.Name).Duplicates()
                .ForEach(name => validation.AppendLine($"The name {name} appears in more than one {BexConstants.AggregateLossSetName.ToLower()}"));

            IndividualLossSetModels.Select(model => model.Name).Duplicates()
                .ForEach(name => validation.AppendLine($"The name {name} appears in more than one {BexConstants.IndividualLossSetName.ToLower()}"));
            
            RateChangeSetModels.Select(model => model.Name).Duplicates()
                .ForEach(name => validation.AppendLine($"The name {name} appears in more than one {BexConstants.RateChangeSetName.ToLower()}"));
        }

        private void ValidatePeriodsDoNotOverlap(StringBuilder validation)
        {
            var periodsOverlap = false;
            foreach (var period in ValidPeriods)
            {
                var n1 = ValidPeriods.Count(x => period.StartDate >= x.StartDate && period.StartDate <= x.EndDate);
                var n2 = ValidPeriods.Count(x => period.EndDate >= x.StartDate && period.EndDate <= x.EndDate);
                if (n1 > 1 || n2 > 1)
                {
                    periodsOverlap = true;
                    break;
                }
            }
            if (periodsOverlap) validation.AppendLine($"{BexConstants.PeriodSetName.ToStartOfSentence()} can't overlap");
        }

        private void ValidatePeriodEndsDoNotPrecedeStarts(StringBuilder validation)
        {
            foreach (var period in ValidPeriods)
            {
                if (period.StartDate > period.EndDate)
                {
                    validation.AppendLine($"{BexConstants.EndDateName.ToStartOfSentence()} <{period.EndDate:d}> can't precede " +
                                          $"{BexConstants.StartDateName.ToLower()} <{period.StartDate:d}>");
                }

            }
        }

        private void ValidatePeriodStartsPrecedeEvaluations(StringBuilder validation)
        {
            foreach (var period in ValidPeriods)
            {
                if (period.StartDate >= period.EvaluationDate)
                {
                    validation.AppendLine($"{BexConstants.StartDateName.ToStartOfSentence()} <{period.StartDate:d}> must precede " +
                                          $"{BexConstants.EvaluationDateName.ToLower()} <{period.EvaluationDate:d}>");
                }
            }
        }

        private void ValidatePeriodsContainAcceptableDates(StringBuilder validation)
        {
            const string messageSuffix = "doesn't fall within acceptable year range";


            foreach (var period in ValidPeriods)
            {
                if (!period.StartDate.IsWithinAcceptableDateRange())
                {
                    validation.AppendLine($"{BexConstants.StartDateName.ToStartOfSentence()} <{period.StartDate:d}> {messageSuffix}.");
                }

                if (!period.EndDate.IsWithinAcceptableDateRange())
                {
                    validation.AppendLine($"{BexConstants.EndDateName.ToStartOfSentence()} <{period.StartDate:d}> {messageSuffix}.");
                }

                if (!period.EvaluationDate.IsWithinAcceptableDateRange())
                {
                    validation.AppendLine($"{BexConstants.EvaluationDateName.ToStartOfSentence()} <{period.StartDate:d}> {messageSuffix}.");
                }
            }
        }
    }
}