using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities
{
    internal class ExcelSheetChangeEventManager
    {
        internal void MonitorSheetChange(Worksheet worksheet, Range range)
        {
            if (worksheet == null) return;

            var package = worksheet.GetPackage();
            if (package != null)
            {
                ModifyPackage(package, range);
                return;
            }

            var segment = worksheet.GetSegment();
            if (segment == null) return;

            //less obvious points
            //the only worksheet segment range in orphan range is prospective exposure 
            //changing a historical period dirties segment
            //changing a subline profile dirties segment
            var intersectingSegmentRangeNames = GetIntersectingSegmentRangeNames(range, segment);
            foreach (var rangeName in intersectingSegmentRangeNames)
            {
                if (segment.OrphanExcelMatrices.Any(matrix => rangeName.Equals(matrix.RangeName)))
                {
                    segment.IsDirty = true;
                }

                var profile = segment.Profiles.SingleOrDefault(prof => ((BaseProfile)prof).CommonExcelMatrix.RangeName == rangeName);
                if (profile != null)
                {
                    profile.IsDirty = true;
                }

                profile = segment.Profiles.SingleOrDefault(prof => ((IProfileExcelMatrix)((BaseProfile)prof).CommonExcelMatrix).BasisRangeName == rangeName);
                if (profile != null)
                {
                    profile.IsDirty = true;

                    var profileExcelMatrix = (IProfileExcelMatrix)((BaseProfile)profile).CommonExcelMatrix;
                    var profileBasisRange = profileExcelMatrix.GetProfileBasisRange();
                    if (!profileExcelMatrix.ValidateBasis())
                    {
                        CorrectProfileBasisValue(profileBasisRange, profileExcelMatrix.FriendlyName);
                    }
                    ReformatProfileRange(profileExcelMatrix, profileBasisRange);
                }

                var periodSet = segment.PeriodSet;
                if (rangeName.Equals(periodSet.ExcelMatrix.RangeName))
                {
                    periodSet.IsDirty = true;
                    segment.IsDirty = true;
                    segment.ExposureSets.Where(expoSet => expoSet.Ledger.Any()).ForEach(expoSet => ModifyExposureLedger(range, expoSet));
                    segment.AggregateLossSets.Where(lossSet => lossSet.Ledger.Any()).ForEach(lossSet => ModifyAggregateLossLedger(range, lossSet));
                }

                if (rangeName.Contains(ExcelConstants.ThresholdRangeName))
                {
                    var componentId = ExcelMatrix.GetComponentIdFromThresholdRangeName(rangeName);
                    segment.IndividualLossSets.Single(ils => ils.ComponentId == componentId).IsDirty = true;
                }

                var historical = segment.Historicals.SingleOrDefault(h => ((BaseHistorical)h).CommonExcelMatrix.RangeName == rangeName);
                if (historical == null) continue;

                historical.IsDirty = true;

                var exposureSet = segment.ExposureSets.SingleOrDefault(expoSet => expoSet.ExcelMatrix.RangeName == rangeName);
                if (exposureSet != null && exposureSet.Ledger.Any()) ModifyExposureLedger(range, exposureSet);
                
                var aggregateLossSet = segment.AggregateLossSets.SingleOrDefault(loss => loss.ExcelMatrix.RangeName == rangeName);
                if (aggregateLossSet != null && aggregateLossSet.Ledger.Any()) ModifyAggregateLossLedger(range, aggregateLossSet);

                var individualLossSet = segment.IndividualLossSets.SingleOrDefault(loss => loss.ExcelMatrix.RangeName == rangeName);
                if (individualLossSet != null && individualLossSet.Ledger.Any()) ModifyIndividualLossLedger(range, individualLossSet);

                var rateChangeSet = segment.RateChangeSets.SingleOrDefault(loss => loss.ExcelMatrix.RangeName == rangeName);
                if (rateChangeSet != null && rateChangeSet.Ledger.Any()) ModifyRateChangeLedger(range, rateChangeSet);
            }
        }

        private static void ModifyRateChangeLedger(Range range, RateChangeSet rateChangeSet)
        {
            var ledger = rateChangeSet.Ledger;

            var excelMatrix = rateChangeSet.ExcelMatrix;
            var periodTopRow = excelMatrix.GetInputRange().GetTopLeftCell().Row;
            var targetTopRow = range.GetTopLeftCell().Row;
            var targetBottomRow = range.GetBottomLeftCell().Row;

            for (var row = targetTopRow; row <= targetBottomRow; row++)
            {
                ledger[row - periodTopRow].IsDirty = true;
            }
        }

        private static void ModifyIndividualLossLedger(Range range,  IndividualLossSet individualLossSet)
        {
            var ledger = individualLossSet.Ledger;
            
            var excelMatrix = individualLossSet.ExcelMatrix;
            var periodTopRow = excelMatrix.GetInputRange().GetTopLeftCell().Row;
            var targetTopRow = range.GetTopLeftCell().Row;
            var targetBottomRow = range.GetBottomLeftCell().Row;

            for (var row = targetTopRow; row <= targetBottomRow; row++)
            {
                ledger[row - periodTopRow].IsDirty = true;
            }
        }

        private static void ModifyExposureLedger(Range range, ExposureSet exposureSet)
        {
            var ledger = exposureSet.Ledger;

            var excelMatrix = exposureSet.ExcelMatrix;
            var periodTopRow = excelMatrix.GetInputRange().GetTopLeftCell().Row;
            var targetTopRow = range.GetTopLeftCell().Row;
            var targetBottomRow = range.GetBottomLeftCell().Row;

            for (var row = targetTopRow; row <= targetBottomRow; row++)
            {
                ledger[row - periodTopRow].IsDirty = true;
            }
        }

        private static void ModifyAggregateLossLedger(Range range, AggregateLossSet aggregateLossSet)
        {
            var ledger = aggregateLossSet.Ledger;

            var excelMatrix = aggregateLossSet.ExcelMatrix;
            var periodTopRow = excelMatrix.GetInputRange().GetTopLeftCell().Row;
            var targetTopRow = range.GetTopLeftCell().Row;
            var targetBottomRow = range.GetBottomLeftCell().Row;

            for (var row = targetTopRow; row <= targetBottomRow; row++)
            {
                ledger[row - periodTopRow].IsDirty = true;
            }
        }

        private static void CorrectProfileBasisValue(Range profileBasisRange, string profileName)
        {
            var incorrectValue = profileBasisRange.Value2;
            
            var up = UserPreferences.ReadFromFile();
            var profileBasisId = up.ProfileBasisId;
            var profileBasis = ProfileBasisFromBex.ReferenceData.Single(x => x.Id == profileBasisId).Name;
            
            var messageStart = incorrectValue == null 
                ? $"Invalid entry: Can't delete {profileName.ToLower()} basis range value." 
                : $"Invalid entry: {profileName.ToStartOfSentence()} basis value <{incorrectValue.ToString()}> is not recognized.";

            var message = $"{messageStart}  Use the dropdown to change this range value. Range value will now get set to {profileBasis.ToLower()}.";

            MessageHelper.Show(message, MessageType.Stop);
            profileBasisRange.Value2 = profileBasis;
        }

        private static void ReformatProfileRange(IProfileExcelMatrix profileExcelMatrix, Range profileBasisRange)
        {
            var profileBasisId = ProfileBasisFromBex.ReferenceData.Single(x => x.Name.Equals(profileBasisRange.GetTopLeftCell().Value2)).Id;
            profileExcelMatrix.ProfileFormatter = ProfileFormatterFactory.Create(profileBasisId);
            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    profileExcelMatrix.Reformat();
                }
            }
        }

        private static IEnumerable<string> GetIntersectingSegmentRangeNames(Range range, ISegment segment)
        {
            var rangeNames = segment.ExcelMatrices.Select(x => x.RangeName).ToList();
            var basisRangeNames = segment.ExcelMatrices.OfType<IProfileExcelMatrix>().Select(rng => rng.BasisRangeName);
            rangeNames.AddRange(basisRangeNames);

            var potentialThresholdRangeNames = segment.IndividualLossSets.Select(ils => ExcelMatrix.GetThresholdRangeName(segment.Id, ils.ComponentId));
            rangeNames.AddRange(potentialThresholdRangeNames.Where(potentialThresholdRangeName => potentialThresholdRangeName.ExistsInWorkbook()));

            if (!segment.IsUmbrella || !segment.ContainsAnyCommercialSublines)
            {
                rangeNames.Remove(segment.UmbrellaExcelMatrix.RangeName);
                rangeNames.Remove(segment.UmbrellaExcelMatrix.BasisRangeName);
            }

            var intersectingRangeNames = range.GetIntersectingRangeNames(rangeNames);
            return intersectingRangeNames;
        }

        private static void ModifyPackage(IPackage package, Range range)
        {
            var packageRangeNames = package.ExcelMatrices.Select(x => x.RangeName);
            if (range.IntersectsAnyNamedRanges(packageRangeNames))
            {
                package.IsDirty = true;
            }
        }
    }
}
