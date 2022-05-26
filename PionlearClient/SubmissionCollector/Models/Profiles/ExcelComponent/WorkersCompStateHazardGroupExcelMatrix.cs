using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public sealed class WorkersCompStateHazardGroupExcelMatrix : SingleOccurrenceProfileExcelMatrix
    {
        public WorkersCompStateHazardGroupExcelMatrix(int segmentId) : base(segmentId)
        {
            InterDisplayOrder = 38;
        }

        public override string FriendlyName => BexConstants.WorkersCompStateHazardGroupProfileName;
        public override string ExcelRangeName => ExcelConstants.WorkersCompStateHazardProfileRangeName;

        [JsonIgnore]
        internal const int ColumnLabelCount = 4;

        [JsonIgnore]
        internal const int RowLabelCount = 3;

        [JsonIgnore]
        public IList<WorkersCompStateHazardGroupAndWeightPlus> Items { get; set; }

        public bool IsIndependent { get; set; }

        internal Range GetHazardGroupRange()
        {
            return GetBodyHeaderRange().GetLastRow().GetRangeSubset(0, ColumnLabelCount);
        }

        internal Range GetHazardGroupPercentRange()
        {
            return GetHazardGroupRange().Offset[-1, 0];
        }

        internal Range GetHazardGroupPremiumTotalRange()
        {
            return GetHazardGroupRange().Offset[-2, 0];
        }

        internal Range GetStatePremiumsRange()
        {
            return GetInputLabelRange().GetRangeSubset(0, 2).GetFirstColumn();
        }

        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = BexConstants.WorkersCompStateHazardGroupProfileName;

            var basisRange = GetProfileBasisRange();
            basisRange.Locked = true;
            basisRange.SetInputLabelColor();
            basisRange.Value = ProfileBasisFromBex.ReferenceData.Single(data => data.Id == BexConstants.PremiumProfileBasisId).Name;
            
            var bodyRange = GetBodyRange();
            bodyRange.ClearBorders();
            bodyRange.SetBorderToOrdinary();

            var bodyHeaderRange = GetBodyHeaderRange();
            bodyHeaderRange.ClearBorders();
            bodyHeaderRange.Locked = true;
            bodyHeaderRange.SetInputLabelColor();
            bodyHeaderRange.SetBorderToOrdinary();

            var hazardGroupPremiumTotalRange = GetHazardGroupPremiumTotalRange();
            hazardGroupPremiumTotalRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            
            var hazardGroupPremiumPercentRange = GetHazardGroupPercentRange();
            hazardGroupPremiumPercentRange.NumberFormat = FormatExtensions.PercentFormat;

            bodyHeaderRange.AlignLeft();
            bodyHeaderRange.GetBottomLeftCell().Resize[1, 4].Value = new[] {"Abbrev", "Name", "Premium", "Premium %"};
            bodyHeaderRange.GetRangeSubset(0, 2).AlignRight();
            bodyHeaderRange.GetColumn(3).SetBorderRightToOrdinary();

            var labelRange = GetInputLabelRange();
            labelRange.SetInputLabelFormat();
            labelRange.AlignLeft();

            var statePremiumsRange = GetStatePremiumsRange();
            statePremiumsRange.AlignRight();
            statePremiumsRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            
            var premiumPercentRange = statePremiumsRange.Offset[0,1];
            premiumPercentRange.AlignRight();
            premiumPercentRange.NumberFormat = FormatExtensions.PercentFormat;
            premiumPercentRange.SetBorderRightToOrdinary();

            var inputRange = GetInputRange();
            inputRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            inputRange.AlignRight();
            
            var premiumTotalFormula = $"=SUM({statePremiumsRange.GetTopLeftCell().Offset[0, 2].Resize[1, 7].Address[false, false]})";
            var hazardGroupPremiumPercentFormula = $"=IFERROR({hazardGroupPremiumTotalRange.GetTopLeftCell().Address[false, false]}/" +
                                                                                   $"SUM({hazardGroupPremiumTotalRange.Address[true, true]}), 0)";

            var segment = GetSegment();
            if (segment.IsWorkerCompClassCodeActive)
            {
                inputRange.SetInputLabelFormat();
                WriteInClassCodeFormula();
                statePremiumsRange.Formula = premiumTotalFormula;
                statePremiumsRange.SetInputLabelFormat();
                hazardGroupPremiumPercentRange.Formula = hazardGroupPremiumPercentFormula;
                hazardGroupPremiumPercentRange.SetInputLabelFormat();
                inputRange.DeEmphasizeZero();
            }
            else
            {
                if (IsIndependent)
                {
                    inputRange.SetInputLabelFormat();
                    var topLeftCell = inputRange.GetTopLeftCell();
                    var statePremiumAddress = GetStatePremiumsRange().GetTopLeftCell().Address[false, true];
                    var hazardGroupAddress = topLeftCell.Offset[-2, 0].Address[true, false];
                    inputRange.Formula = $"=IF( OR(ISBLANK({statePremiumAddress}), ISBLANK({hazardGroupAddress})), \"\", {statePremiumAddress} * {hazardGroupAddress})";

                    statePremiumsRange.SetInputFormat();
                    hazardGroupPremiumPercentRange.SetInputFormat();
                    inputRange.DeEmphasizeZero();
                }
                else
                {
                    inputRange.SetInputFormat();
                    statePremiumsRange.Formula = premiumTotalFormula;
                    statePremiumsRange.SetInputLabelFormat();
                    hazardGroupPremiumPercentRange.Formula = hazardGroupPremiumPercentFormula;
                    hazardGroupPremiumPercentRange.SetInputLabelFormat();
                    inputRange.ClearConditionalFormats();
                }
            }

            hazardGroupPremiumTotalRange.Formula = $"=SUM({inputRange.GetFirstColumn().Address[false, false]})"; 
            premiumPercentRange.Formula = $"=IFERROR({statePremiumsRange.GetTopLeftCell().Address[false, false]}/" +
                                               $"SUM({statePremiumsRange.Address}), 0)";

            bodyRange.ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;
        }

        public override void ImplementProfileBasis()
        {
            
        }

        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(RowLabelCount, ColumnLabelCount);
        }

        public override Range GetInputLabelRange()
        {
            return RangeName.GetRangeSubset(RowLabelCount, 0).GetFirstColumns(ColumnLabelCount);
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetFirstRows(RowLabelCount);
        }
        
        public override StringBuilder Validate()
        {
            Items = new List<WorkersCompStateHazardGroupAndWeightPlus>(); 
            
            var validation = new StringBuilder();
            var segment = GetSegment();
            if (segment.IsWorkerCompClassCodeActive) return validation;
            
            if (IsIndependent)
            {
                ValidateIndependent(validation);
            }
            else
            {
                ValidateOrdinary(validation);
            }


            return validation;
        }

        private void ValidateOrdinary(StringBuilder validation)
        {
            var gridRange = GetInputRange();
            var gridFromExcel = gridRange.GetContent();
            if (gridFromExcel.AllNull())
            {
                return;
            }

            var gridAsDouble = gridFromExcel.ForceContentToDoubles();

            var hazardGroups = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.OrderBy(hg => hg.DisplayOrder).ToList();
            var columnCount = hazardGroups.Count;
            var stateIds = GetStateIds();
            var rowCount = stateIds.Count;

            var startColumn = gridRange.GetTopLeftCell().Column;
            var startRow = gridRange.GetTopLeftCell().Row;
            var columnLetters = GetColumnLetters(columnCount, startColumn);

            for (var row = 0; row < rowCount; row++)
            {
                var rowNumber = row + startRow;
                var stateId = stateIds[row];

                for (var column = 0; column < columnCount; column++)
                {
                    var gridItemFromExcel = gridFromExcel[row, column];
                    if (gridItemFromExcel == null) continue;

                    var gridItem = gridAsDouble[row, column];
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetters[column], rowNumber);

                    if (double.IsNaN(gridItem))
                    {
                        validation.AppendLine($"Enter value in {addressLocation}:" +
                                              $" <{gridItemFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                    }

                    Items.Add(new WorkersCompStateHazardGroupAndWeightPlus
                    {
                        WorkersCompStateId = stateId,
                        WorkersCompHazardGroupId = hazardGroups[column].Id,
                        Value = gridItem,
                        Location = addressLocation
                    });
                }
            }
        }

        private void ValidateIndependent(StringBuilder validation)
        {
            var hazardGroupPercentRange = GetHazardGroupPercentRange();
            var statePremiumsRange = GetStatePremiumsRange();

            var isHazardGroupEmpty = hazardGroupPercentRange.IsEmpty();
            var isPremiumEmpty = statePremiumsRange.IsEmpty();

            if (isHazardGroupEmpty && isPremiumEmpty)
            {
                return;
            }

            if (isHazardGroupEmpty ^ isPremiumEmpty)
            {
                validation.AppendLine("Only one of the input ranges has content");
                return;
            }

            
            var hazardGroupContent = hazardGroupPercentRange.GetContent().ForceContentToDoubles().GetRow(0).ToList().Where(item => !double.IsNaN(item)).ToList();
            if (hazardGroupContent.Any(item => item > 1))
            {
                validation.AppendLine($"Hazard group percents need to < {1.00:P1}");
            }
            var sum = hazardGroupContent.Sum();
            if (!sum.IsEpsilonEqual(1d, .01))
            {
                validation.AppendLine($"Hazard group percents need to sum to 100% (between {0.99:P1} and {1.01:P1}).  The sum is currently {sum:P1}");
            }
            
            var gridRange = GetInputRange();
            var gridFromExcel = gridRange.GetContent();
            var gridAsDouble = gridFromExcel.ForceContentToDoubles();

            var hazardGroupPercentsForExcel = hazardGroupPercentRange.GetContent();
            var hazardGroupPercentsAsDouble = hazardGroupPercentsForExcel.ForceContentToDoubles();

            var statePremiumsFromExcel = statePremiumsRange.GetContent();
            var statePremiumsAsDouble = statePremiumsFromExcel.ForceContentToDoubles();

            var hazardGroups = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.OrderBy(hg => hg.DisplayOrder).ToList();
            var columnCount = hazardGroups.Count;
            var stateIds = GetStateIds();
            var rowCount = stateIds.Count;

            var startColumn = gridRange.GetTopLeftCell().Column;
            var startRow = gridRange.GetTopLeftCell().Row;
            var columnLetters = GetColumnLetters(columnCount, startColumn);

            var addressLocation = string.Empty;
            for (var row = 0; row < rowCount; row++)
            {
                var rowNumber = row + startRow;
                var stateId = stateIds[row];

                var statePremiumFromExcel = statePremiumsFromExcel[row, 0];
                var statePremiumAsDouble = statePremiumsAsDouble[row, 0];

                if (statePremiumFromExcel != null && double.IsNaN(statePremiumAsDouble))
                {
                    var columnLetter = GetColumnLetters(1, startColumn-2).First().Value;
                    addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber); 
                    validation.AppendLine($"Enter state premium in row {rowNumber}:" + $" <{statePremiumFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }

                for (var column = 0; column < columnCount; column++)
                {
                    var hazardGroupPercentFromExcel = hazardGroupPercentsForExcel[0, column];
                    var hazardGroupPercentAsDouble = hazardGroupPercentsAsDouble[0, column];

                    if (row == 0 && hazardGroupPercentFromExcel != null && double.IsNaN(hazardGroupPercentAsDouble))
                    {
                        addressLocation = RangeExtensions.GetAddressLocation(columnLetters[column], startRow - 2);
                        validation.AppendLine($"Enter hazard group percent in {addressLocation}:" +
                                              $" <{hazardGroupPercentFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                    }

                    var gridItemFromExcel = gridFromExcel[row, column];
                    if (gridItemFromExcel == null) continue;

                    var gridItem = gridAsDouble[row, column];
                    if (double.IsNaN(gridItem)) continue;
                    
                    Items.Add(new WorkersCompStateHazardGroupAndWeightPlus
                    {
                        WorkersCompStateId = stateId,
                        WorkersCompHazardGroupId = hazardGroups[column].Id,
                        Value = gridItem,
                        Location = addressLocation
                    });
                }
            }
        }

        public static string GetRangeName(int segmentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.WorkersCompStateHazardProfileRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId)
        {
            return $"{GetRangeName(segmentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetBasisRangeName(int segmentId)
        {
            return $"{GetRangeName(segmentId)}.{ExcelConstants.ProfileBasisRangeName}";
        }

        
        public override void ToggleEstimate()
        {
            if (GetSegment().IsWorkerCompClassCodeActive)
            {
                MessageHelper.Show("Can't toggle estimate on protected range", MessageType.Stop);
                return;
            }

            base.ToggleEstimate();
        }

        public override Range GetBodyRange()
        {
            return RangeName.GetRangeSubset(RowLabelCount, 0);
        }

        internal void WriteInClassCodeFormula()
        {
            var segment = GetSegment();

            //when user preferences has class codes checked, first reformat call can't link to a yet-to-exist range
            var classCodeProfile = segment.WorkersCompClassCodeProfile;
            var excelMatrix = classCodeProfile?.ExcelMatrix;
            if (excelMatrix == null) return;
            if (!excelMatrix.RangeName.ExistsInWorkbook()) return;

            var range = GetInputRange();
            var topLeftRange = range.GetTopLeftCell();
          
            var classCodeRange = excelMatrix.GetInputRange();
            var classCodeStateAbbreviationRange = classCodeRange.GetColumn(0); 
            var classCodePremiumRange = classCodeRange.GetColumn(2);
            var classCodeHazardGroupRange = classCodeRange.GetColumn(3);

            var formula = "=SUMIFS(" +
                          $"{classCodePremiumRange.Address[true, true]}, " +
                          $"{classCodeStateAbbreviationRange.Address[true, true]}, " +
                          $"{topLeftRange.Offset[0, -ColumnLabelCount].Address[false, true]}, " +
                          $"{classCodeHazardGroupRange.Address[true, true]}, " +
                          $"{topLeftRange.Offset[-1, 0].Address[true, false]}" +
                          ")";
            range.Formula = formula;
        }

        private static Dictionary<int, string> GetColumnLetters(int columnCount, int startColumn)
        {
            var columnLetters = new Dictionary<int, string>();
            for (var column = 0; column < columnCount; column++)
            {
                columnLetters.Add(column, (startColumn + column).GetColumnLetter());
            }

            return columnLetters;
        }
        
        private List<int> GetStateIds()
        {
            var stateLabelRange = GetInputLabelRange();
            var stateAbbreviations = stateLabelRange.GetContent().ForceContentToStrings().GetColumn(0).ToList();
            var states = StateCodesFromBex.GetWorkersCompStates().ToDictionary(key => key.Abbreviation);
            var stateIds = new List<int>();
            foreach (var stateAbbreviation in stateAbbreviations)
            {
                stateIds.Add(states[stateAbbreviation].Id);
            }

            return stateIds;
        }

    }
}
