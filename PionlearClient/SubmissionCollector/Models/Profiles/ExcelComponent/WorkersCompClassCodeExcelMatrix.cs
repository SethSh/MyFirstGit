using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public sealed class WorkersCompClassCodeExcelMatrix : SingleOccurrenceProfileExcelMatrix, IRangeResizable
    {
        public WorkersCompClassCodeExcelMatrix(int segmentId) : base(segmentId)
        {
            InterDisplayOrder = 37;
        }

        public override string FriendlyName => BexConstants.WorkersCompClassCodeName;
        public override string ExcelRangeName => ExcelConstants.WorkersCompClassCodeProfileRangeName;

        internal const int BufferRowCount = 3;
        public Range GetInputRangeWithoutBuffer()
        {
            return GetInputRange().RemoveLastRows(BufferRowCount); 
        }

        public Range GetPureInputRangeWithoutBuffer()
        {
            const int pureInputColumnCount = 3;
            return GetInputRangeWithoutBuffer().GetFirstColumns(pureInputColumnCount);
        }

        [JsonIgnore]
        public IList<WorkersCompStateClassCodeAndValuePlus> Items { get; set; }


        public override void Reformat()
        {
            base.Reformat();

            const int stateAbbreviationColumnIndex = 0;
            const int stateCodeColumnIndex = 1;
            const int premiumColumnIndex = 2;
            const int labelColumnOffset = premiumColumnIndex+1;
            
            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = BexConstants.WorkersCompClassCodeProfileName;

            var bodyRange = GetBodyRange();
            bodyRange.ClearBorders();
            bodyRange.SetBorderToOrdinary();

            var bodyHeaderRange = GetBodyHeaderRange();
            bodyHeaderRange.ClearBorders();
            bodyHeaderRange.GetTopLeftCell().Locked = true;
            bodyHeaderRange.GetTopLeftCell().SetInputLabelColor();
            bodyHeaderRange.SetBorderToOrdinary();
            bodyHeaderRange.AlignLeft();
            bodyHeaderRange.GetColumn(premiumColumnIndex).AlignRight();
            bodyHeaderRange.Value = new [] {"State Abbrev", "Code", "Premium", "Hazard Group", "Class Code Desc"};
            bodyHeaderRange.GetColumn(stateCodeColumnIndex).AlignRight();
            bodyHeaderRange.GetColumn(premiumColumnIndex).AlignRight();
            bodyHeaderRange.Locked = true;
            bodyHeaderRange.SetInputLabelColor();

            var inputRangeWithoutBuffer = GetInputRangeWithoutBuffer();
            inputRangeWithoutBuffer.SetBorderAroundToResizable();
            inputRangeWithoutBuffer.AlignLeft();
            inputRangeWithoutBuffer.GetColumn(stateAbbreviationColumnIndex).NumberFormat = FormatExtensions.StringFormat;
            inputRangeWithoutBuffer.GetColumn(stateCodeColumnIndex).NumberFormat = FormatExtensions.WholeNumberWithFourLeadingZerosFormat;
            inputRangeWithoutBuffer.GetColumn(stateCodeColumnIndex).AlignRight();
            inputRangeWithoutBuffer.GetColumn(premiumColumnIndex).NumberFormat = FormatExtensions.WholeNumberFormat;
            inputRangeWithoutBuffer.GetColumn(premiumColumnIndex).AlignRight();
            inputRangeWithoutBuffer.GetRangeSubset(0,labelColumnOffset).NumberFormat = FormatExtensions.StringFormat;


            if (GetSegment().IsWorkerCompClassCodeActive)
            {
                inputRangeWithoutBuffer.SetInputFormat();
                inputRangeWithoutBuffer.GetFirstColumn().SetInputDropdownInteriorColor();
                inputRangeWithoutBuffer.GetRangeSubset(0, labelColumnOffset).SetInputLabelFormat();

                var stateHazardGroupExcelMatrix = GetSegment().WorkersCompStateHazardGroupProfile.ExcelMatrix;
                stateHazardGroupExcelMatrix.WriteInClassCodeFormula();
            }
            else
            {
                inputRangeWithoutBuffer.SetInputLabelFormat();
            }

            var inputRangeBuffer = GetInputRange().GetLastRows(BufferRowCount);
            inputRangeBuffer.ClearContents();
            inputRangeBuffer.Locked = true;
            inputRangeBuffer.SetInputLabelColor();
            inputRangeBuffer.ClearBorderAllButTop();
            inputRangeBuffer.AlignRight();
            inputRangeBuffer.GetColumn(0).AlignLeft();
            WriteInTotals(premiumColumnIndex, inputRangeWithoutBuffer, inputRangeBuffer);

            bodyRange.GetColumn(stateAbbreviationColumnIndex).ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetColumn(stateCodeColumnIndex).ColumnWidth = ExcelConstants.LessThanStandardColumnWidth;
            bodyRange.GetColumn(premiumColumnIndex).ColumnWidth = ExcelConstants.StandardColumnWidth; 
            bodyRange.GetColumn(premiumColumnIndex+1).ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetLastColumn().ColumnWidth = ExcelConstants.ExtraLargeColumnWidth;
            bodyRange.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;
        }
        

        public override void ImplementProfileBasis()
        {
            
        }

        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public override Range GetInputLabelRange()
        {
            throw new ArgumentException();
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }
        
        public override StringBuilder Validate()
        {
            const int stateColumnIndex = 0;
            const int classCodeColumnIndex = stateColumnIndex + 1;
            const int premiumColumnIndex = stateColumnIndex + 2;

            Items = new List<WorkersCompStateClassCodeAndValuePlus>(); 
            
            var validation = new StringBuilder();
            if (!GetSegment().IsWorkerCompClassCodeActive) return validation;

            var gridRange = GetInputRangeWithoutBuffer();
            var gridFromExcel = gridRange.GetContent();
            if (gridFromExcel.AllNull()) return validation;

            var stateAbbreviationsFromExcel = gridFromExcel.GetColumn(stateColumnIndex);
            var classCodesFromExcel = gridFromExcel.GetColumn(classCodeColumnIndex);
            var premiumsFromExcel = gridFromExcel.GetColumn(premiumColumnIndex);

            var stateAbbreviations = stateAbbreviationsFromExcel.ForceContentToStrings();
            var classCodes = classCodesFromExcel.ForceContentToStrings()
                .Select(item => long.TryParse(item, out var classAsLong) ? classAsLong : new long?()).ToArray();
            var premiums = premiumsFromExcel.ForceContentToNullableDoubles();

            var topLeftCell = gridRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var rowCount = gridFromExcel.GetLength(0);

            var columnLetters = startColumn.GetColumnLetters(premiumColumnIndex + 1);
            var stateClassCodeSets = WorkersCompClassCodesAndHazardsFromBex.StateClassCodes.ToList();
            var referenceStateClassCodes = stateClassCodeSets.ToDictionary(key => key.State.Id, value => value.ClassCodes);
            var stateAbbreviationMap = stateClassCodeSets.ToDictionary(codeSet => codeSet.State.Abbreviation, codeSet => codeSet.State.Id);
            
            for (var row = 0; row < rowCount; row++)
            {
                var rowNumber = startRow + row;
                
                var stateAbbreviationFromExcel = stateAbbreviationsFromExcel[row];
                var classCodeFromExcel = classCodesFromExcel[row];
                var premiumFromExcel = premiumsFromExcel[row];

                if (stateAbbreviationFromExcel == null && classCodeFromExcel == null && premiumFromExcel == null) continue;

                var stateAbbreviation = stateAbbreviations[row];
                var classCode = classCodes[row];
                var premium = premiums[row];


                var stateId = new int?();
                var stateLocation = RangeExtensions.GetAddressLocation(columnLetters[startColumn + stateColumnIndex], rowNumber);
                if (stateAbbreviationFromExcel == null)
                {
                    validation.AppendLine($"Enter state in {stateLocation}");
                }
                else 
                {
                    if (stateAbbreviationMap.Keys.Contains(stateAbbreviation)) stateId = stateAbbreviationMap[stateAbbreviation];
                    if (!stateId.HasValue)
                    {
                        validation.AppendLine($"Enter valid state in {stateLocation}: <{stateAbbreviationFromExcel}> {BexConstants.NotRecognizedAsAState}");
                    }
                }

                var hasHazard = false; 
                var classCodeId = new long?();
                var classCodeLocation = RangeExtensions.GetAddressLocation(columnLetters[startColumn + classCodeColumnIndex], rowNumber);
                if (classCodeFromExcel == null)
                {
                    validation.AppendLine($"Enter class code in {classCodeLocation}");
                }
                else
                {
                    if (!classCode.HasValue)
                    {
                        validation.AppendLine($"Enter class code in {classCodeLocation}: <{classCodesFromExcel}> <{BexConstants.NotRecognizedAsANumber}>");
                    }
                    else if (stateId.HasValue)
                    {
                        var referenceClassCodes = referenceStateClassCodes[stateId.Value];
                        var referenceClassCode = referenceClassCodes.SingleOrDefault(item => item.StateClassCode == classCode);
                        if (referenceClassCode == null)
                        {
                            classCodeId = classCode;
                        }
                        else
                        {
                            classCodeId = referenceClassCode.Id;
                            //we had a plan to allow unmapped class codes ... now we decided against it
                            //these won't go to BEX
                            hasHazard = referenceClassCode.HazardGroupId.HasValue;
                        }
                    }
                }
                

                var premiumLocation = RangeExtensions.GetAddressLocation(columnLetters[startColumn + premiumColumnIndex], rowNumber);
                if (premiumFromExcel == null)
                {
                    validation.AppendLine($"Enter premium in {premiumLocation}");
                }
                else if (premium.HasValue && double.IsNaN(premium.Value))
                {
                    validation.AppendLine($"Enter premium in {premiumLocation}: <{premiumFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }

                if (validation.Length == 0)
                {
                    Debug.Assert(stateId != null, nameof(stateId) + " != null");
                    Debug.Assert(classCodeId != null, nameof(classCodeId) + " != null");
                    Debug.Assert(premium != null, nameof(premium) + " != null");

                    Items.Add(new WorkersCompStateClassCodeAndValuePlus
                    {
                        WorkersCompStateId = stateId.Value,
                        ClassCodeId = classCodeId.Value,
                        Value = premium.Value,
                        Location = $"row {rowNumber}",
                        RowNumber = rowNumber,
                        HasHazard = hasHazard
                    });
                }
            }

            return validation;
        }

       
        public static string GetRangeName(int segmentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.WorkersCompClassCodeProfileRangeName}";
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
            IsEstimate = !IsEstimate;
            var inputRange = GetInputRangeWithoutBuffer();

            SetInputDropdownInteriorColorContemplatingEstimate(inputRange.GetFirstColumn());
            SetInputInteriorColorContemplatingEstimate(inputRange.GetRangeSubset(0,1).GetFirstColumns(2));
        }
        
        public override Range GetBodyRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public void ReformatBorderTop()
        {
            GetInputRange().GetFirstRow().SetBorderTopToOrdinary();
        }

        public void ReformatBorderBottom()
        {
            RangeName.GetRange().GetLastRow().SetBorderBottomToResizable();
        }

        public override void MoveRangesWhenSublinesChange(int sublineCount)
        {
            if (!RangeName.ExistsInWorkbook()) return;
            base.MoveRangesWhenSublinesChange(sublineCount);
        }

        private static void WriteInTotals(int premiumColumnIndex, Range inputRangeWithoutBuffer, Range inputRangeBuffer)
        {
            var hazardGroupColumnIndex = premiumColumnIndex + 1;
            var premiumAddress = inputRangeWithoutBuffer.GetColumn(premiumColumnIndex).Address;
            var hazardGroupAddress = inputRangeWithoutBuffer.GetColumn(hazardGroupColumnIndex).Address;

            var anchor = inputRangeBuffer.GetTopLeftCell();
            anchor.Value2 = "Total Premium";
            anchor.Offset[1, 0].Value2 = "Mapped Premium";
            anchor.Offset[2, 0].Value2 = "Un-Mapped Premium";

            anchor.Offset[0, premiumColumnIndex].Formula = $"=SUM({premiumAddress})";
            anchor.Offset[1, premiumColumnIndex].Formula = $"=SUMIF({hazardGroupAddress}, \"<>\", {premiumAddress})";
            anchor.Offset[2, premiumColumnIndex].Formula = $"=SUMIF({hazardGroupAddress}, \"\", {premiumAddress})";

            var premiumSumAddress = anchor.Offset[0, premiumColumnIndex].Address;
            var thisRange = anchor.Offset[1, hazardGroupColumnIndex];
            thisRange.Formula = $"=IFERROR({thisRange.Offset[0, -1].Address} / {premiumSumAddress}, 0)";
            thisRange = anchor.Offset[2, hazardGroupColumnIndex];
            thisRange.Formula = $"=IFERROR({thisRange.Offset[0, -1].Address} / {premiumSumAddress}, 0)";

            inputRangeBuffer.GetColumn(premiumColumnIndex).NumberFormat = FormatExtensions.WholeNumberFormat;
            inputRangeBuffer.GetColumn(hazardGroupColumnIndex).NumberFormat = FormatExtensions.PercentFormatWithoutDecimal;
        }
    }
}
