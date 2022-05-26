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
    public sealed class AggregateLossSetExcelMatrix : MultipleOccurrenceSegmentLossExcelMatrix, IRangeSumable, IRangeResizable
    {
        public AggregateLossSetExcelMatrix(int segmentId, int componentId) : base(segmentId)
        {
            ComponentId = componentId;
            InterDisplayOrder = 50;
        }

        public AggregateLossSetExcelMatrix(int segmentId) : base(segmentId)
        {

        }

        public AggregateLossSetExcelMatrix()
        {
                
        }

        public override string FriendlyName => BexConstants.AggregateLossSetName;
        public override string ExcelRangeName => ExcelConstants.AggregateLossSetRangeName;
        [JsonIgnore] public override bool HasEmptyColumnToRight => false;

        [JsonIgnore]
        public List<AggregateLossModelPlus> Items { get; set; }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maximumDisplayOrder = GetSegment().AggregateLossSets.Max(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != maximumDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {FriendlyName.ToLower()} farther to the right.", MessageType.Stop);
                return false;
            }
        }

        public override bool IsOkToMoveLeft
        {
            get
            {
                var minimumDisplayOrder = GetSegment().AggregateLossSets.Min(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != minimumDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {FriendlyName.ToLower()} farther to the left.",
                    MessageType.Stop);
                return false;
            }
        }

        public override void MoveRight()
        {
            var segment = GetSegment();
            var excelMatrices = segment.AggregateLossSets.Select(a => a.ExcelMatrix).ToList();

            var excelComponentToRight = excelMatrices.Single(h => h.IntraDisplayOrder == IntraDisplayOrder + 1);
            var isRightest = excelComponentToRight.IntraDisplayOrder == excelMatrices.Max(matrix => matrix.IntraDisplayOrder);
            var columnShift = excelComponentToRight.RangeName.GetRange().Columns.Count;

            var range = RangeName.GetRange();
            var columnCount = range.Columns.Count;

            range.Offset[0, columnCount + columnShift].InsertColumnsToRight();
            range.EntireColumn.Cut(range.Offset[0, columnCount + columnShift].EntireColumn);
            range.Offset[0, -(columnCount + columnShift)].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);

            if (!isRightest) return;

            RangeName.GetRange().SetBorderRightToResizable();
            HeaderRangeName.GetRange().SetBorderRightToOrdinary();
            excelComponentToRight.RangeName.GetRange().SetBorderRightToOrdinary();
        }

        public override void MoveLeft(int steps = 1)
        {
            //when called from ribbon display order is ok to use
            //when called from subline wizard, display order is not ok to use because model and worksheet are out of sync
            var segment = GetSegment();
            var siblingExcelMatrices = segment.AggregateLossSets.Select(a => a.ExcelMatrix).ToList();

            var columnStarts = siblingExcelMatrices.Select(p => p.ColumnStart).ToList();
            columnStarts.Sort();

            var otherColumnStart = columnStarts[columnStarts.IndexOf(ColumnStart) - steps];
            var columnShift = ColumnStart - otherColumnStart;
            var siblingExcelComponent = siblingExcelMatrices.Single(s => s.ColumnStart == otherColumnStart);
            var isRightest = ColumnStart == columnStarts.Last();

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

            r = siblingExcelComponent.RangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            r = siblingExcelComponent.HeaderRangeName.GetRange();
            r.SetBorderRightToOrdinary();
            r.SetBorderLeftToOrdinary();

            if (!isRightest) return;
            
            columnStarts = siblingExcelMatrices.Select(p => p.ColumnStart).ToList();
            columnStarts.Sort();
            siblingExcelMatrices.Single(s => s.ColumnStart == columnStarts.Last()).RangeName.GetRange().SetBorderRightToResizable();
        }

        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.AggregateLossSetRangeName}{componentId}";
        }
        
        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return GetHeaderRangeName(segmentId, componentId, ExcelConstants.AggregateLossSetRangeName);
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.SublinesRangeName}";
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return $"{GetSublinesRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public void ReformatBorderTop()
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
            sublinesHeaderRange.GetTopLeftCell().Value = $"{BexConstants.AggregateLossSetShortName} {BexConstants.SublineName}s";

            var sublinesRange = GetSublinesRange();
            sublinesRange.SetSublineFormat(); 
            
            var columnCount = sublinesRange.Columns.Count;
            for (var rowIndex = 0; rowIndex < sublinesRange.Rows.Count; rowIndex++)
            {
                if (sublinesRange.Offset[rowIndex, 0].GetTopLeftCell().Value2 != null)
                {
                    sublinesRange.Offset[rowIndex, 0].Resize[1, columnCount].AlignCenterAcrossSelection();
                }
                else
                {
                    sublinesRange.Offset[rowIndex, 0].Resize[1, columnCount].AlignLeft();
                }
            }

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.GetTopLeftCell().Value = BexConstants.AggregateLossSetShortName;

            var inputLabelRange = GetInputLabelRange();
            inputLabelRange.SetBorderAroundToOrdinary();
            inputLabelRange.SetInputLabelColor();
            inputLabelRange.Locked = true;
            var labels = GetColumnLabels(); 
            inputLabelRange.Value = labels.ToOneByNArray();
            inputLabelRange.AlignRight();

            var inputRange = GetInputRange();
            inputRange.AlignRight();
            inputRange.SetBorderAroundToResizable();
            inputRange.NumberFormat = FormatExtensions.WholeNumberFormat;

            
            var sumRange = GetSumRange();
            sumRange.AlignRight();
            sumRange.Formula = $"=Sum({inputRange.GetFirstColumn().Address[false, false]})";
            sumRange.ClearBorderAllButTop();
            sumRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            sumRange.SetInputLabelInteriorColor();
            sumRange.SetInputLabelFontColor();
            sumRange.Locked = true;

            inputRange.ColumnWidth = ExcelConstants.LossColumnWidth;
        }

        private IList<string> GetColumnLabels()
        {
            var list = new List<string>();
            var desc = GetSegment().AggregateLossSetDescriptor;
            if (desc.IsPaidAvailable && desc.IsLossAndAlaeCombined) 
            {
                list.Add(BexConstants.PaidLossAndAlaeName);
            }
            else if (desc.IsPaidAvailable && !desc.IsLossAndAlaeCombined)
            {
                list.Add(BexConstants.PaidLossName);
                list.Add(BexConstants.PaidAlaeName);
            }

            if (desc.IsLossAndAlaeCombined)
            {
                list.Add(BexConstants.ReportedLossAndAlaeName);
            }
            else
            {
                list.Add(BexConstants.ReportedLossName);
                list.Add(BexConstants.ReportedAlaeName);
            }
            return list;
        }

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
            var periodSet = GetSegment().PeriodSet;
            var periodsValidationDetails = periodSet.ExcelMatrix.ValidationDetails;

            var validation = new StringBuilder();
            Items = new List<AggregateLossModelPlus>();

            var inputRange = GetInputRange();

            var amountsFromExcel = inputRange.GetContent();
            var amounts = amountsFromExcel.ForceContentToDoubles();

            var descriptor = GetSegment().AggregateLossSetDescriptor;
            var isLossAndAlaeCombined = descriptor.IsLossAndAlaeCombined;
            var isLossAndAlaeSeparate = !descriptor.IsLossAndAlaeCombined;
            var isPaidAvailable = descriptor.IsPaidAvailable;

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var lossSet = (AggregateLossSet)GetParent();
            var ledger = lossSet.Ledger;

            var dataSetId = GetParent().SourceId;

            var rowCount = amountsFromExcel.GetLength(0);
            if (!ledger.Any()) ledger.Create(rowCount);

            var aggPaidLossRowCount = 0;
            var aggReportedLossRowCount = 0;

            for (var row = 0; row < rowCount; row++)
            {
                var rowNumber = row + startRow;

                var periodValidationDetail = periodsValidationDetails[row];
                var rowFromExcel = amountsFromExcel.GetRow(row);
                var rowAsDouble = amounts.GetRow(row);

                var columnIndex = 0;
                var paidLoss = isPaidAvailable && isLossAndAlaeSeparate ? rowFromExcel[columnIndex++] : null;
                var paidAlae = isPaidAvailable && isLossAndAlaeSeparate ? rowFromExcel[columnIndex++] : null;
                var paidLossAndAlae = isPaidAvailable && isLossAndAlaeCombined ? rowFromExcel[columnIndex++] : null;

                var reportedLoss = isLossAndAlaeSeparate ? rowFromExcel[columnIndex++] : null;
                var reportedAlae = isLossAndAlaeSeparate ? rowFromExcel[columnIndex++] : null;
                var reportedLossAndAlae = isLossAndAlaeCombined ? rowFromExcel[columnIndex] : null;
                
                columnIndex = 0;
                var paidLossAsDouble = isPaidAvailable && isLossAndAlaeSeparate ? rowAsDouble[columnIndex++] : double.NaN;
                var paidAlaeAsDouble = isPaidAvailable && isLossAndAlaeSeparate ? rowAsDouble[columnIndex++] : double.NaN;
                var paidLossAndAlaeAsDouble = isPaidAvailable && isLossAndAlaeCombined ? rowAsDouble[columnIndex++] : double.NaN;

                var reportedLossAsDouble = isLossAndAlaeSeparate ? rowAsDouble[columnIndex++] : double.NaN;
                var reportedAlaeAsDouble = isLossAndAlaeSeparate ? rowAsDouble[columnIndex++] : double.NaN;
                var reportedLossAndAlaeAsDouble = isLossAndAlaeCombined ? rowAsDouble[columnIndex] : double.NaN;

                const int lossAndAlaeColumn = 0;
                const int lossColumn = 0;
                const int alaeColumn = 1;

                var paidColumnLetters = new List<string>();
                var reportedColumnLetters = new List<string>();
                if (isPaidAvailable)
                {
                    paidColumnLetters.Add(startColumn.GetColumnLetter());
                    if (isLossAndAlaeSeparate)
                    {
                        paidColumnLetters.Add((startColumn + 1).GetColumnLetter());
                        reportedColumnLetters.Add((startColumn + 2).GetColumnLetter());
                        reportedColumnLetters.Add((startColumn + 3).GetColumnLetter());
                    }
                    else
                    {
                        reportedColumnLetters.Add((startColumn + 1).GetColumnLetter());
                    }
                }
                else
                {
                    reportedColumnLetters.Add(startColumn.GetColumnLetter());
                    if (isLossAndAlaeSeparate)
                    {
                        reportedColumnLetters.Add((startColumn + 1).GetColumnLetter());
                    }
                }
                

                if (rowFromExcel.AllNull() && periodValidationDetail.IsNull)
                {
                    continue;
                }

                if (rowAsDouble.Any(amount => amount.IsNotNaN()) && periodValidationDetail.IsNull)
                {
                    var location = RangeExtensions.GetAddressLocation(periodValidationDetail.ColumnLetters, rowNumber);
                    validation.AppendLine($"Enter the {BexConstants.PeriodName.ToLower()} in {location}");
                    continue;
                }

                if (rowFromExcel.AllNull() && periodValidationDetail.IsOk)
                {
                    continue;
                }

                if (isLossAndAlaeCombined)
                {
                    if (isPaidAvailable)
                    {
                        if (paidLossAndAlae != null && double.IsNaN(paidLossAndAlaeAsDouble))
                        {
                            var addressLocation = RangeExtensions.GetAddressLocation(paidColumnLetters[lossAndAlaeColumn], rowNumber);
                            var amountFromExcel = rowFromExcel[lossAndAlaeColumn];
                            validation.AppendLine(
                                $"{BexConstants.PaidLossAndAlaeName.ToStartOfSentence()} <{amountFromExcel}> in {addressLocation} " +
                                $"{BexConstants.NotRecognizedAsANumber}");
                        }

                        if (paidLossAndAlae != null) aggPaidLossRowCount++;
                    }

                    

                    if (reportedLossAndAlae != null && double.IsNaN(reportedLossAndAlaeAsDouble))
                    {
                        var addressLocation = RangeExtensions.GetAddressLocation(reportedColumnLetters[lossAndAlaeColumn], rowNumber); 
                        var amountFromExcel = rowFromExcel[lossAndAlaeColumn];
                        validation.AppendLine(
                            $"{BexConstants.ReportedLossAndAlaeName.ToStartOfSentence()} <{amountFromExcel}> in " +
                            $"{addressLocation} {BexConstants.NotRecognizedAsANumber}");
                    }

                    if (reportedLossAndAlae != null) aggReportedLossRowCount++;
                }
                else
                {
                    if (isPaidAvailable)
                    {
                        
                        if (paidLoss != null && double.IsNaN(paidLossAsDouble))
                        {
                            var paidLossAddressLocation = RangeExtensions.GetAddressLocation(paidColumnLetters[lossColumn], rowNumber);
                            var amountFromExcel = rowFromExcel[lossColumn];
                            validation.AppendLine(
                                $"{BexConstants.PaidLossName.ToStartOfSentence()} <{amountFromExcel}> in {paidLossAddressLocation} " +
                                $"{BexConstants.NotRecognizedAsANumber}");
                        }

                        if (paidAlae != null && double.IsNaN(paidAlaeAsDouble))
                        {
                            var paidAlaeAddressLocation = RangeExtensions.GetAddressLocation(paidColumnLetters[alaeColumn], rowNumber);
                            var amountFromExcel = rowFromExcel[alaeColumn];
                            validation.AppendLine(
                                $"{BexConstants.PaidAlaeName.ToStartOfSentence()} <{amountFromExcel}> in {paidAlaeAddressLocation} " +
                                $"{BexConstants.NotRecognizedAsANumber}");
                        }

                        if (!(paidLoss == null && paidAlae == null)) aggPaidLossRowCount++;
                    }
                    
                    if (reportedLoss != null && double.IsNaN(reportedLossAsDouble))
                    {
                        var reportedLossAddressLocation = RangeExtensions.GetAddressLocation(reportedColumnLetters[lossColumn], rowNumber);
                        var amountFromExcel = rowFromExcel[lossColumn];
                        validation.AppendLine(
                            $"{BexConstants.ReportedLossName.ToStartOfSentence()} <{amountFromExcel}> in {reportedLossAddressLocation} " +
                            $"{BexConstants.NotRecognizedAsANumber}");
                    }

                    if (reportedAlae != null && double.IsNaN(reportedAlaeAsDouble))
                    {
                        var reportedAlaeAddressLocation = RangeExtensions.GetAddressLocation(reportedColumnLetters[alaeColumn], rowNumber); 
                        var amountFromExcel = rowFromExcel[alaeColumn];
                        validation.AppendLine(
                            $"{BexConstants.ReportedAlaeName.ToStartOfSentence()} <{amountFromExcel}> in {reportedAlaeAddressLocation} " +
                            $"{BexConstants.NotRecognizedAsANumber}");
                    }

                    if (!(reportedLoss == null && reportedAlae == null)) aggReportedLossRowCount++;
                }

                var item = new AggregateLossModelPlus
                {
                    RowId = row,
                    RowNumber = rowNumber,
                    SourceId = ledger[row].SourceId,
                    IsDirty = ledger[row].IsDirty,

                    DataSetId = dataSetId,

                    StartDate = periodValidationDetail.StartDate.MapToOffset(),
                    EndDate = periodValidationDetail.EndDate.MapToOffset(),
                    
                    PaidLossAmount = paidLoss != null ? paidLossAsDouble : new double?(),
                    PaidAlaeAmount = paidAlae != null ? paidAlaeAsDouble : new double?(),
                    PaidCombinedAmount = paidLossAndAlae != null ? paidLossAndAlaeAsDouble : new double?(),

                    ReportedLossAmount = reportedLoss != null ? reportedLossAsDouble : new double?(),
                    ReportedAlaeAmount = reportedAlae!= null ?  reportedAlaeAsDouble : new double?(),
                    ReportedCombinedAmount = reportedLossAndAlae != null ? reportedLossAndAlaeAsDouble : new double?(),
                };

                Items.Add(item);
            }

            if (aggReportedLossRowCount != 0 && aggReportedLossRowCount < periodSet.ExcelMatrix.Items.Count)
            {
                validation.AppendLine(isLossAndAlaeCombined
                    ? $"{BexConstants.ReportedLossAndAlaeName.ToStartOfSentence()} must be all supplied or all blank"
                    : $"At least one of {BexConstants.ReportedLossName.ToLower()} and {BexConstants.ReportedAlaeName.ToLower()} " +
                      "must be supplied for all periods or must all be blank");
            }

            if (aggPaidLossRowCount != 0 && aggPaidLossRowCount < periodSet.ExcelMatrix.Items.Count)
            {
                validation.AppendLine(isLossAndAlaeCombined
                    ? $"{BexConstants.PaidLossAndAlaeName.ToStartOfSentence()} must be all supplied or all blank"
                    : $"At least one of {BexConstants.PaidLossName.ToLower()} and {BexConstants.PaidAlaeName.ToLower()} " +
                      "must be supplied for all periods or must all be blank");
            }

            return validation;
        }
        
        public IModel GetParent()
        {
            return GetSegment().AggregateLossSets.Single(x => x.ComponentId == ComponentId);
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetFirstRow();
        }

        public override Range GetBodyRange()
        {
            return GetInputRange();
        }

        public Range GetSumRange()
        {
            return RangeName.GetRange().GetLastRow().Offset[1, 0];
        }

        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var aggregateLossSets = GetSegment().AggregateLossSets;
            var otherExcelMatrix = aggregateLossSets.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1).ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var aggregateLossSets = GetSegment().AggregateLossSets;
            var otherExcelMatrix = aggregateLossSets.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1).ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }

        protected override void InsertPaidColumns()
        {
            if (DoLabelsContainAnExactMatch(new List<string>{BexConstants.PaidLossName, BexConstants.PaidLossAndAlaeName})) return;

            var labelRange = GetInputLabelRange();
            var reportedColumnIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.ReportedName).ToList();
            var additionalColumnCount = reportedColumnIndices.Count;

            var firstReportedColumnIndex = reportedColumnIndices.First();
            labelRange.Offset[0, firstReportedColumnIndex].Resize[1, additionalColumnCount].InsertColumnsToRight();

            GetSublinesHeaderRange().AppendColumnsToLeft(additionalColumnCount).SetInvisibleRangeName(SublinesHeaderRangeName);
            GetSublinesRange().AppendColumnsToLeft(additionalColumnCount).SetInvisibleRangeName(SublinesRangeName);
            GetHeaderRange().AppendColumnsToLeft(additionalColumnCount).SetInvisibleRangeName(HeaderRangeName);
            var range = RangeName.GetRange();
            range.AppendColumnsToLeft(additionalColumnCount).SetInvisibleRangeName(RangeName);
            range.ColumnWidth = range.GetTopRightCell().ColumnWidth;

            labelRange = GetInputLabelRange();
            for (var counter = 0; counter < additionalColumnCount; counter++)
            {
                labelRange.GetTopLeftCell().Offset[0, counter].Value = labelRange.GetTopLeftCell()
                    .Offset[0, additionalColumnCount + counter]
                    .Value.ToString().Replace(BexConstants.ReportedName, BexConstants.PaidName);
            }

            var headerRange = GetHeaderRange();
            var sublinesHeaderRange = GetSublinesHeaderRange();
            var sublinesRange = GetSublinesRange();

            headerRange.GetTopLeftCell().Offset[0, additionalColumnCount].MoveRangeContent(headerRange.GetTopLeftCell());
            sublinesHeaderRange.GetTopLeftCell().Offset[0, additionalColumnCount].MoveRangeContent(sublinesHeaderRange.GetTopLeftCell());
            sublinesRange.GetFirstColumn().Offset[0, additionalColumnCount].MoveRangeContent(sublinesRange.GetFirstColumn());
        }

        protected override void DeletePaidColumns()
        {
            if (!DoLabelsContainAnExactMatch(new List<string> {BexConstants.PaidLossName, BexConstants.PaidLossAndAlaeName})) return; 
            
            var labelRange = GetInputLabelRange();
            var paidColumnIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.PaidName).ToList();
            var paidColumnCount = paidColumnIndices.Count;

            var headerRange = GetHeaderRange();
            var sublinesHeaderRange = GetSublinesHeaderRange();
            var sublinesRange = GetSublinesRange();

            headerRange.GetTopLeftCell().MoveRangeContent(headerRange.GetTopLeftCell().Offset[0, paidColumnCount]);
            sublinesHeaderRange.GetTopLeftCell().MoveRangeContent(sublinesHeaderRange.GetTopLeftCell().Offset[0, paidColumnCount]);
            sublinesRange.GetFirstColumn().MoveRangeContent(sublinesRange.GetFirstColumn().Offset[0, paidColumnCount]);

            paidColumnIndices.Reverse();
            foreach (var index in paidColumnIndices)
            {
                labelRange.Resize[1, 1].Offset[0, index].EntireColumn.Delete();
            }
        }
    }
}
