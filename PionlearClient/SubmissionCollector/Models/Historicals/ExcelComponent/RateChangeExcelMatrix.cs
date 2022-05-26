using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.CollectorClientPlus;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelWorkspaceFolder;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.View.Forms;
using IModel = PionlearClient.Model.IModel;

namespace SubmissionCollector.Models.Historicals.ExcelComponent
{
    [JsonObject]
    public sealed class RateChangeExcelMatrix: MultipleOccurrenceSegmentExcelMatrix, IRangeResizable, IRangeFillable
    {
        public RateChangeExcelMatrix(int segmentId, int componentId) : base(segmentId)
        {
            ComponentId = componentId;
            InterDisplayOrder = 65;
        }

        public RateChangeExcelMatrix(int segmentId) : base(segmentId)
        {

        }
        
        public RateChangeExcelMatrix()
        {
                
        }

        [JsonIgnore]
        public List<RateChangeModelPlus> Items { get; set; }

        public override string FriendlyName => BexConstants.RateChangeSetName;
        public override string ExcelRangeName => ExcelConstants.RateChangeSetRangeName;


        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public override Range GetInputLabelRange()
        {
            return RangeName.GetRange().GetFirstRow();
        }

        public override StringBuilder Validate()
        {
            var validation = new StringBuilder();
            Items = new List<RateChangeModelPlus>();

            var inputRange = GetInputRange();

            var datesFromExcel = inputRange.GetColumn(0).GetContent();
            var dates = datesFromExcel.ForceContentToNullableDates();
            var datesColumnLetter = inputRange.GetTopLeftCell().Column.GetColumnLetter();

            var ratesFromExcel = inputRange.GetColumn(1).GetContent();
            var rates = ratesFromExcel.ForceContentToNullableDoubles();
            var ratesColumnLetter = inputRange.GetTopLeftCell().Offset[0, 1].Column.GetColumnLetter();
            
            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            
            var rateChangeSet = (RateChangeSet) GetParent();
            var ledger = rateChangeSet.Ledger;

            var dataSetId = GetParent().SourceId;

            var rowCount = datesFromExcel.GetLength(0);
            if (!ledger.Any()) ledger.Create(rowCount);

            for (var row = 0; row < rowCount; row++)
            {
                var rowNumber = row + startRow;

                var dateFromExcel = datesFromExcel[row, 0];
                var rateFromExcel = ratesFromExcel[row, 0];
            
                if (dateFromExcel == null && rateFromExcel == null)
                {
                    continue;
                }

                var date = dates[row, 0];
                var rate = rates[row, 0];
                var hasIssue = false;

                if (!date.HasValue)
                {
                    var location = RangeExtensions.GetAddressLocation(datesColumnLetter, rowNumber);
                    if (dateFromExcel == null)
                    {
                        validation.AppendLine($"Enter the {BexConstants.RateChangeName.ToLower()} date in {location}");
                    }
                    else
                    {
                        validation.AppendLine($"{BexConstants.RateChangeName.ToStartOfSentence()} " +
                                              $"date <{dateFromExcel}> in {location} is not recognized as a date");
                    }
                    hasIssue = true;
                }

                if (!rate.HasValue)
                {
                    var location = RangeExtensions.GetAddressLocation(ratesColumnLetter, rowNumber);
                    validation.AppendLine($"Enter the {BexConstants.RateChangeName.ToLower()} rate in {location}");
                    hasIssue = true;
                }
                else if (double.IsNaN(rate.Value))
                {
                    var location = RangeExtensions.GetAddressLocation(ratesColumnLetter, rowNumber);
                    validation.AppendLine($"{BexConstants.RateChangeName.ToStartOfSentence()} " +
                                          $"rate <{rateFromExcel}> in {location} is not recognized as a number");
                    hasIssue = true;
                }
                
                
                if (hasIssue) continue;
                var item = new RateChangeModelPlus
                {
                    RowId = row,
                    RowNumber = rowNumber,
                    SourceId = ledger[row].SourceId,
                    IsDirty = ledger[row].IsDirty,

                    DataSetId = dataSetId,

                    EffectiveDate = date.Value.MapToOffset(),
                    Value = rate.Value,
                };

                Items.Add(item);
            }

            return validation;
        }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maximumDisplayOrder = GetSegment().RateChangeSets.Max(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != maximumDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {FriendlyName.ToLower()} farther to the right.", MessageType.Stop);
                return false;
            }
        }

        public override bool IsOkToMoveLeft {
            get
            {
                var minimumDisplayOrder = GetSegment().RateChangeSets.Min(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != minimumDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {FriendlyName.ToLower()} farther to the left.", MessageType.Stop);
                return false;
            }
        }
        public override void MoveRight()
        {
            var segment = GetSegment();
            var siblingExcelMatrices = segment.RateChangeSets.Select(a => a.ExcelMatrix).ToList();

            var excelMatrix = siblingExcelMatrices.Single(h => h.IntraDisplayOrder == IntraDisplayOrder + 1);
            var columnShift = excelMatrix.RangeName.GetRange().AppendColumn().Columns.Count;

            var range = RangeName.GetRange().AppendColumn();
            var columnCount = range.Columns.Count;

            range.Offset[0, columnCount + columnShift].InsertColumnsToRight();
            range.EntireColumn.Cut(range.Offset[0, columnCount + columnShift].EntireColumn);
            range.Offset[0, -(columnCount + columnShift)].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        public override void MoveLeft(int steps = 1)
        {
            //when called from ribbon display order is ok to use
            //when called from subline wizard, display order is not ok to use because model and worksheet are out of sync
            var segment = GetSegment();
            var siblingExcelMatrices = segment.RateChangeSets.Select(a => a.ExcelMatrix).ToList();

            var columnStarts = siblingExcelMatrices.Select(p => p.ColumnStart).ToList();
            columnStarts.Sort();

            var myColumnStart = ColumnStart;
            var otherColumnStart = columnStarts[columnStarts.IndexOf(myColumnStart) - steps];
            var columnShift = myColumnStart - otherColumnStart;

            var range = RangeName.GetRange().AppendColumn();
            var columnCount = range.Columns.Count;

            range.Offset[0, -columnShift].InsertColumnsToRight();
            range.EntireColumn.Cut(range.Offset[0, -(columnShift + columnCount)].EntireColumn);
            range.Offset[0, columnShift + columnCount].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().RateChangeSets
                .Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1)
                .ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().RateChangeSets.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1)
                .ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }

        public IModel GetParent()
        {
            return GetSegment().RateChangeSets.Single(x => x.ComponentId == ComponentId);
        }

        public override void Reformat()
        {
            base.Reformat();

            var sublinesHeaderRange = GetSublinesHeaderRange();
            sublinesHeaderRange.SetSublineHeaderFormat();
            sublinesHeaderRange.ClearContents();
            sublinesHeaderRange.GetTopLeftCell().Value = $"{BexConstants.RateChangeSetName} {BexConstants.SublineName}";
            sublinesHeaderRange.GetTopRightCell().Offset[0, 1].ColumnWidth = FormatExtensions.ColumnWidthEmptyDefault;

            var sublinesRange = GetSublinesRange();
            sublinesRange.SetSublineFormat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = BexConstants.RateChangeSetName;

            var inputLabelRange = GetInputLabelRange();
            inputLabelRange.SetInputLabelFormatWithBorder();
            inputLabelRange.ClearContents();
            inputLabelRange.AlignRight();
            inputLabelRange.GetTopLeftCell().Value = "Date";
            inputLabelRange.GetTopLeftCell().Offset[0,1].Value = "Rate Change";
            
            var inputRange = GetInputRange();
            inputRange.Locked = false;
            inputRange.SetBorderAroundToResizable();
            inputRange.SetInputColor();
            inputRange.GetFirstColumn().FormatWithDates();
            inputRange.GetColumn(1).FormatWithPercents();

            var bufferRange = inputRange.GetLastRow().Offset[1, 0];
            bufferRange.SetInputLabelFormat();
            bufferRange.ClearBorderAllButTop();
        }

        public void ReformatBorderTop()
        {
            RangeName.GetRange().GetFirstRow().SetBorderTopToOrdinary();
        }

        public void ReformatBorderBottom()
        {
            RangeName.GetRange().GetLastRow().SetBorderBottomToOrdinary();
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }

        public override Range GetBodyRange()
        {
            return GetInputRange();
        }

        public static string GetRangeHeaderName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.RateChangeSetRangeName}{componentId}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return GetHeaderRangeName(segmentId, componentId, ExcelConstants.RateChangeSetRangeName);
        }
        
        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return $"{GetSublinesRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.SublinesRangeName}";
        }
    }
}
