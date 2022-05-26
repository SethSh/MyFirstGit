using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.View;
using SubmissionCollector.View.Forms;
using IModel = PionlearClient.Model.IModel;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public sealed class TotalInsuredValueExcelMatrix : MultipleOccurrenceProfileExcelMatrix, IRangeResizable
    {
        public TotalInsuredValueExcelMatrix(int segmentId) : base(segmentId)
        {
            InterDisplayOrder = 36;
        }

        public TotalInsuredValueExcelMatrix()
        {

        }

        public override string FriendlyName => BexConstants.TotalInsuredValueProfileName;
        public override string ExcelRangeName => ExcelConstants.TotalInsuredValueProfileRangeName;

        public ISubline Subline => this.Single();

        public override int ComponentId
        {
            get => this.Single().Code;
            set { }
        }

        [JsonIgnore]
        public IList<PionlearClient.CollectorClientPlus.TotalInsuredValueDistributionItemPlus> Items { get; set; }

        [JsonIgnore] 
        public bool IsExpanded => ((TotalInsuredValueProfile) GetParent()).IsExpanded;


        public override void ModifyForChangeInSublines(int sublineCount)
        {
            var segment = GetSegment();
            var multipleOccurrenceSegmentExcelMatrices = segment.ExcelMatrices.OfType<MultipleOccurrenceSegmentExcelMatrix>().ToList();
            var maximumSublineCount = multipleOccurrenceSegmentExcelMatrices.Max(c => c.Count);
            MoveRangesWhenSublinesChange(maximumSublineCount);
        }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maxDisplayOrder = GetSegment().TotalInsuredValueProfiles.Max(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != maxDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {FriendlyName.ToLower()} farther to the right.", MessageType.Stop);
                return false;
            }
        }
        public override bool IsOkToMoveLeft
        {
            get
            {
                var minDisplayOrder = GetSegment().TotalInsuredValueProfiles.Min(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != minDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {FriendlyName.ToLower()} farther to the left.",
                    MessageType.Stop);
                return false;
            }
        }

        
        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = this.GetParent().Name;
            headerRange.SetToDefaultFont();

            var basisRange = GetProfileBasisRange();
            basisRange.Locked = false;
            basisRange.SetInputDropdownInteriorColor();
            basisRange.SetInputFontColor();
            SetProfileBasesInWorksheet();

            var bodyRange = GetBodyRange();
            bodyRange.ClearBorders();
            
            var bodyHeaderRange = GetBodyHeaderRange();
            bodyHeaderRange.ClearBorders();
            bodyHeaderRange.GetTopLeftCell().Locked = true;
            bodyHeaderRange.GetTopLeftCell().SetInputLabelColor();
            bodyHeaderRange.SetBorderToOrdinary();
            bodyHeaderRange.GetTopLeftCell().AlignRight();
            if (IsExpanded)
            {
                bodyHeaderRange.GetTopLeftCell().Resize[1, 4].Value = new [] {BexConstants.TivAverageName, BexConstants.ShareName, BexConstants.LimitName, BexConstants.SirAttachmentName };
            }
            else
            {
                bodyHeaderRange.GetTopLeftCell().Value = BexConstants.TivAverageName;
            }
            
            bodyHeaderRange.GetTopRightCell().AlignRight();
            bodyHeaderRange.SetBodyHeaderFormat();

            var inputRange = GetInputRange();
            var inputRangeWithoutBuffer = inputRange.RemoveLastRow();
            inputRangeWithoutBuffer.AlignRight();
            inputRangeWithoutBuffer.SetInputColor();
            inputRangeWithoutBuffer.SetBorderAroundToResizable();

            var inputRangeBuffer = inputRange.GetLastRow();
            inputRangeBuffer.ClearContents();
            
            var profile = GetSegment().TotalInsuredValueProfiles.Single(x => x.ComponentId == ComponentId);
            var containsCommercial = profile.ExcelMatrix.ContainsCommercial();
            var note = containsCommercial
                ? "Note: TIVs include Buildings and Contents" 
                : "Note: TIVs include Buildings Only";
            inputRangeBuffer.GetTopLeftCell().Value = note; 
            
            inputRangeBuffer.Locked = true;
            inputRangeBuffer.SetInputLabelColor();

            bodyRange.GetColumn(0).ColumnWidth = ExcelConstants.StandardColumnWidth; 
            bodyRange.GetColumn(0).FormatWithWholeNumbers();
            if (IsExpanded)
            {
                bodyRange.GetColumn(1).ColumnWidth = ExcelConstants.StandardColumnWidth;
                bodyRange.GetColumn(1).FormatWithPercents();
                bodyRange.GetColumn(2).ColumnWidth = ExcelConstants.StandardColumnWidth;
                bodyRange.GetColumn(2).FormatWithWholeNumbers();
                bodyRange.GetColumn(3).ColumnWidth = ExcelConstants.StandardColumnWidth;
                bodyRange.GetColumn(3).FormatWithWholeNumbers();
            }
            bodyRange.GetLastColumn().ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetLastColumn().FormatWithPercents();

            bodyRange.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;

            ImplementProfileBasis();
        }

        public override void ImplementProfileBasis()
        {
            var weightRange = GetInputRange().GetLastColumn();
            var sumRange = GetSumRange();
            var sumFormula = $"=Sum({weightRange.Address})";

            ProfileFormatter.FormatDataRange(weightRange);
            ProfileFormatter.WriteSumFormulaToRange(sumRange, sumFormula);
        }

        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public Range GetInputRangeWithoutBuffer()
        {
            return GetInputRange().RemoveLastRow();
        }

        public override Range GetInputLabelRange()
        {
            return RangeName.GetRange().GetFirstRow();
        }

        public override StringBuilder Validate()
        {
            var validation = new StringBuilder();
            Items = new List<PionlearClient.CollectorClientPlus.TotalInsuredValueDistributionItemPlus>();

            var inputRange = GetInputRangeWithoutBuffer();
            var columnCount = inputRange.Columns.Count;
            
            var gridFromExcel = inputRange.GetContent();
            var gridAsDouble = gridFromExcel.ForceContentToDoubles();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            
            var isExpanded = IsExpanded;
            var weightLabelFromExcel = GetProfileBasisRange().Value2;
            var weightLabel = weightLabelFromExcel != null ? weightLabelFromExcel.ToString() : string.Empty;

            var rowCount = gridAsDouble.GetLength(0);
            for (var row = 0; row < rowCount; row++)
            {
                var insuredValue = gridFromExcel[row, 0];
                var value = gridFromExcel[row, columnCount-1];
                var share = isExpanded ? gridFromExcel[row, 1] : null; 
                var limit = isExpanded ? gridFromExcel[row, 2] : null;
                var sir = isExpanded ? gridFromExcel[row, 3] : null;

                if (insuredValue == null && share == null && limit == null && sir == null && value == null) continue;

                var insuredValueDouble = gridAsDouble[row, 0];
                var valueDouble = gridAsDouble[row, columnCount - 1];
                var shareDouble = isExpanded ? gridAsDouble[row, 1] : double.NaN; 
                var limitDouble = isExpanded ? gridAsDouble[row, 2] : double.NaN;
                var sirDouble = isExpanded ? gridAsDouble[row, 3] : double.NaN;

                var absoluteRow = row + startRow;

                var blankErrors = GetBlankErrors(insuredValue, value, weightLabel);
                var dataTypeErrors = GetDataTypeErrors(insuredValueDouble, insuredValue, 
                    share, shareDouble, 
                    limit, limitDouble, 
                    sir, sirDouble, 
                    valueDouble, value, weightLabel);

                var rowPrefix = $"Row {absoluteRow}:";
                if (dataTypeErrors.Count > 0)
                {
                    var dataTypeCsv = string.Join(", ", dataTypeErrors);
                    var dataTypeMessage = dataTypeErrors.Count > 0 ? $"{dataTypeCsv} {BexConstants.NotRecognizedAsANumber}" : string.Empty;
                    
                    validation.AppendLine($"{rowPrefix} {dataTypeMessage}");
                }

                if (blankErrors.Count > 0)
                {
                    var blankCsv = string.Join(", ", blankErrors);
                    var blankMessage = blankErrors.Count > 0 ? $"{blankCsv} is blank" : string.Empty;

                    validation.AppendLine($"{rowPrefix} {blankMessage}");
                    continue;
                }

                if (dataTypeErrors.Count > 0 || blankErrors.Count > 0)
                {
                    continue;
                }

                var item = new PionlearClient.CollectorClientPlus.TotalInsuredValueDistributionItemPlus
                {
                    TotalInsuredValue = insuredValueDouble,
                    Limit = limit == null ? new double?() : limitDouble,
                    Attachment = sir == null ? new double?() : sirDouble,
                    Share = share == null ? new double?() : shareDouble,
                    Weight = valueDouble,
                    Location = $"row {absoluteRow}"
                };
                Items.Add(item);
            }

            return validation;
        }

        public IModel GetParent()
        {
            return GetSegment().TotalInsuredValueProfiles.Single(x => x.ComponentId == ComponentId);
        }

        
        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.TotalInsuredValueProfileRangeName}{componentId}";
        }


        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().TotalInsuredValueProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1).ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().TotalInsuredValueProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1).ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }

        public override void ToggleEstimate()
        {
            IsEstimate = !IsEstimate;

            var inputRange = GetInputRangeWithoutBuffer();
            SetInputInteriorColorContemplatingEstimate(inputRange);
        }

        
        public static string GetBasisRangeName(int segmentId, int componentId)
        {
            return $"segment{segmentId}.{ExcelConstants.TotalInsuredValueProfileRangeName}{componentId}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return GetHeaderRangeName(segmentId, componentId, ExcelConstants.TotalInsuredValueProfileRangeName);
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return GetSublinesRangeName(segmentId, componentId, ExcelConstants.TotalInsuredValueProfileRangeName);
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return GetSublinesHeaderRangeName(segmentId, componentId, ExcelConstants.TotalInsuredValueProfileRangeName);
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
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


        public void SynchronizeExpansion()
        {
            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    if (IsExpanded)
                    {
                        ExpandedColumns();
                    }
                    else
                    {
                        UnExpandedColumns();
                    }
                    Reformat();
                }
            }
        }

        private void ExpandedColumns()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.ShareName)) return;

            var range = GetInputRange();
            range.Offset[0, 1].Resize[1, 3].InsertColumnsToRight();
        }

        private void  UnExpandedColumns()
        {
            if (!DoLabelsContainAnExactMatch(BexConstants.ShareName)) return;

            var range = GetInputRange();
            range.Offset[0, 1].Resize[1, 3].EntireColumn.Delete();
        }

        private string GetProfileName()
        {
            return IsExpanded ? BexConstants.TotalInsuredValueProfileName : BexConstants.TotalInsuredValueAbbreviatedProfileName;
        }

        private static List<string> GetBlankErrors(object insuredValue, object value, string weightLabel)
        {
            var errors = new List<string>();
            if (insuredValue == null)
            {
                errors.Add("TIV");
            }

            
            if (value == null)
            {
                errors.Add(weightLabel);
            }

            return errors;
        }

        private static List<string> GetDataTypeErrors(double insuredValueDouble, object insuredValue, object share, double shareDouble, object limit,
            double limitDouble, object sir, double sirDouble, double valueDouble, object value, string weightLabel )
        {
            var dataTypeErrors = new List<string>();
            if (insuredValue != null && double.IsNaN(insuredValueDouble))
            {
                dataTypeErrors.Add($"TIV <{insuredValue}>");
            }

            if (share != null && double.IsNaN(shareDouble))
            {
                dataTypeErrors.Add($"Share <{share}>");
            }

            if (limit != null && double.IsNaN(limitDouble))
            {
                dataTypeErrors.Add($"Limit <{limit}>");
            }

            if (sir != null && double.IsNaN(sirDouble))
            {
                dataTypeErrors.Add($"SIR <{sir}>");
            }

            if (value != null && double.IsNaN(valueDouble))
            {
                dataTypeErrors.Add($"{weightLabel} <{value}>");
            }

            return dataTypeErrors;
        }

    }
}
