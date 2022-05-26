using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.CollectorClientPlus;
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
    public sealed class ExposureSetExcelMatrix : MultipleOccurrenceSegmentExcelMatrix, IRangeSumable, IRangeResizable
    {
        public ExposureSetExcelMatrix(int segmentId, int componentId) : base(segmentId)
        {
            ComponentId = componentId;
            InterDisplayOrder = 40;
        }

        public ExposureSetExcelMatrix()
        {
        
        }

        public override string FriendlyName => BexConstants.ExposureSetName;
        public override string ExcelRangeName => ExcelConstants.ExposureSetRangeName;
        [JsonIgnore] public override bool HasEmptyColumnToRight => false;

        [JsonIgnore]
        public List<ExposureModelPlus> Items { get; set; }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maxDisplayOrder = GetSegment().ExposureSets.Max(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != maxDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {BexConstants.ExposureSetName.ToLower()} farther to the right.",
                    MessageType.Stop);
                return false;
            }
        }
        public override bool IsOkToMoveLeft
        {
            get
            {
                var minDisplayOrder = GetSegment().ExposureSets.Min(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != minDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {BexConstants.ExposureSetName.ToLower()} farther to the left.",
                    MessageType.Stop);
                return false;
            }
        }

        public override void MoveRight()
        {
            var segment = GetSegment();
            var siblingExcelMatrices = segment.ExposureSets.Select(a => a.ExcelMatrix).ToList();

            var excelMatrixToRight = siblingExcelMatrices.Single(h => h.IntraDisplayOrder == IntraDisplayOrder + 1);
            var columnShift = excelMatrixToRight.RangeName.GetRange().Columns.Count;

            var range = RangeName.GetRange();
            var columnCount = range.Columns.Count;

            range.Offset[0, columnCount + columnShift].InsertColumnsToRight();
            range.EntireColumn.Cut(range.Offset[0, columnCount + columnShift].EntireColumn);
            range.Offset[0, -(columnCount + columnShift)].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);

            var r = RangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            r = HeaderRangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            r = excelMatrixToRight.RangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            r = excelMatrixToRight.HeaderRangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();
        }

        public override void MoveLeft(int steps = 1)
        {
            //when called from ribbon display order is ok to use
            //when called from subline wizard, display order is not ok to use because model and worksheet are out of sync
            var segment = GetSegment();
            var siblingExcelMatrices = segment.ExposureSets.Select(a => a.ExcelMatrix).ToList();

            var columnStarts = siblingExcelMatrices.Select(p => p.ColumnStart).ToList();
            columnStarts.Sort();

            var myColumnStart = ColumnStart;
            var otherColumnStart = columnStarts[columnStarts.IndexOf(myColumnStart) - steps];
            var columnShift = myColumnStart - otherColumnStart;
            var siblingExcelMatrix = siblingExcelMatrices.Single(s => s.ColumnStart == otherColumnStart);

            var range = RangeName.GetRange();
            var columnCount = range.Columns.Count;

            range.Offset[0, -columnShift].InsertColumnsToRight();
            range.EntireColumn.Cut(range.Offset[0, -(columnShift + columnCount)].EntireColumn);
            range.Offset[0, columnShift + columnCount].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);

            var r = RangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            r = HeaderRangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            r = siblingExcelMatrix.RangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            r = siblingExcelMatrix.HeaderRangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();
        }
        
        public  void ReformatBorderTop()
        {
            GetInputRange().GetFirstRow().SetBorderTopToOrdinary();
        }

        public void ReformatBorderBottom()
        {
            RangeName.GetRange().GetLastRow().SetBorderBottomToResizable();
        }

        public override void Reformat()
        {
            base.Reformat();

            var sublinesHeaderRange = GetSublinesHeaderRange();
            sublinesHeaderRange.SetSublineHeaderFormat();
            sublinesHeaderRange.ClearContents();
            sublinesHeaderRange.GetTopLeftCell().Value = $"{BexConstants.ExposureSetName} {BexConstants.SublineName}s";

            var sublinesRange = GetSublinesRange();
            sublinesRange.SetSublineFormat(); 
            
            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.SetBorderRightToOrdinary();
            headerRange.SetBorderLeftToOrdinary();
            headerRange.GetTopLeftCell().Value = BexConstants.ExposureSetName;

            var basisRange = GetInputLabelRange();
            basisRange.SetBorderRightToOrdinary();
            basisRange.SetBorderLeftToOrdinary();
            basisRange.AlignRight();
            basisRange.Locked = true;
            basisRange.SetInputLabelColor();

            var inputRange = GetInputRange();
            inputRange.SetBorderAroundToResizable();
            inputRange.AlignRight();
            inputRange.NumberFormat = FormatExtensions.WholeNumberFormat;

            var sumRange = GetSumRange();
            sumRange.Formula = $"=Sum({inputRange.Address})";
            sumRange.ClearBorderAllButTop();
            sumRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            sumRange.Locked = true;
            sumRange.SetInputLabelColor();

            inputRange.ColumnWidth = ExcelConstants.LossColumnWidth;
        }
        
        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public override Range GetInputLabelRange()
        {
            return RangeName.GetTopLeftCell();
        }

        public override StringBuilder Validate()
        {
            var periodSet = GetSegment().PeriodSet;
            var periodsValidationDetails = periodSet.ExcelMatrix.ValidationDetails;
            var validation = new StringBuilder();
            Items = new List<ExposureModelPlus>();

            var inputRange = GetInputRange();

            var amountsFromExcel = inputRange.GetContent();
            var amounts = amountsFromExcel.ForceContentToDoubles();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var exposureSet = (ExposureSet)GetParent();
            var ledger = exposureSet.Ledger;

            var dataSetId = GetParent().SourceId;

            const int amountColumn = 0;
            var columnLetter = startColumn.GetColumnLetter();
            var rowCount = amountsFromExcel.GetLength(0);
            if (!ledger.Any()) ledger.Create(rowCount);

            for (var row = 0; row < rowCount ; row++)
            {
                var rowNumber = row + startRow;

                var periodValidationDetail = periodsValidationDetails[row];
                var amountFromExcel = amountsFromExcel[row, amountColumn];
                var amount = amounts[row, amountColumn];

                if (amountFromExcel == null && periodValidationDetail.IsNull)
                {
                    continue;
                }

                if (amount.IsNotNaN() && periodValidationDetail.IsNull)
                {
                    var location = RangeExtensions.GetAddressLocation(periodValidationDetail.ColumnLetters, rowNumber);
                    validation.AppendLine($"Enter {BexConstants.PeriodName.ToLower()} in {location}");
                    continue;
                }

                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                if (amountFromExcel == null && periodValidationDetail.IsOk)
                {
                    validation.AppendLine($"Enter {BexConstants.ExposureName.ToLower()} in {addressLocation}");
                    continue;
                }

                if (double.IsNaN(amount))
                {
                    validation.AppendLine(amountFromExcel == null 
                        ? $"Enter {BexConstants.ExposureName.ToLower()} in {addressLocation}"
                        : $"Enter {BexConstants.ExposureName.ToLower()} in {addressLocation}: <{amountFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                    continue;
                }
                
                var item = new ExposureModelPlus
                {
                    RowId = row,
                    Location = addressLocation,

                    SourceId = ledger[row].SourceId,
                    IsDirty = ledger[row].IsDirty,

                    DataSetId = dataSetId,

                    StartDate = periodValidationDetail.StartDate.MapToOffset(),
                    EndDate = periodValidationDetail.EndDate.MapToOffset(),

                    Amount = amount
                };
                Items.Add(item);
            }

            return validation;
        }

        public IModel GetParent()
        {
            return GetSegment().ExposureSets.Single(x => x.ComponentId == ComponentId);
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }

        public override Range GetBodyRange()
        {
            return GetInputRange();
        }

        public Range GetSumRange()
        {
            return RangeName.GetRange().GetBottomLeftCell().Offset[1, 0];
        }

        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().ExposureSets.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1).ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().ExposureSets.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1).ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }
        
        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.ExposureSetRangeName}{componentId}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return GetHeaderRangeName(segmentId, componentId, ExcelConstants.ExposureSetRangeName);
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.SublinesRangeName}";
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return $"{GetSublinesRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }
    }
}
