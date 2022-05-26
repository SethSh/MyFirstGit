using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
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
    //IndividualLossSetExcelMatrix can't change name bc of existing workbooks
    [JsonObject]
    public sealed class ExcelMatrix : MultipleOccurrenceSegmentLossExcelMatrix, IRangeResizable
    {
        private static readonly List<string> NumberLabels = new List<string>
        {
            BexConstants.LimitName,
            BexConstants.SirAttachmentName,
            BexConstants.LossName,
            BexConstants.AlaeName
        };


        public ExcelMatrix(int segmentId, int componentId) : base(segmentId)
        {
            ComponentId = componentId;
            InterDisplayOrder = 60;
        }

        public ExcelMatrix(int segmentId) : base(segmentId)
        {

        }

        public ExcelMatrix()
        {

        }

        public override string FriendlyName => BexConstants.IndividualLossSetName;
        public override string ExcelRangeName => ExcelConstants.IndividualLossSetRangeName;

        [JsonIgnore]
        public List<IndividualLossModelPlus> Items { get; set; }

        [JsonIgnore] public int? Threshold { get; set; }

        public override void Reformat()
        {
            base.Reformat();

            var sublinesHeaderRange = GetSublinesHeaderRange();
            sublinesHeaderRange.SetSublineHeaderFormat();
            sublinesHeaderRange.ClearContents();
            sublinesHeaderRange.GetTopLeftCell().Value = $"{BexConstants.IndividualLossSetName} {BexConstants.SublineName}s";

            GetSublinesHeaderRange().SetBorderAroundToOrdinary();
            var sublinesRange = GetSublinesRange();
            sublinesRange.SetBorderAroundToOrdinary();
            sublinesRange.AlignCenterAcrossSelection();
            GetInputLabelRange().SetBorderAroundToOrdinary();

            var headerRange = HeaderRangeName.GetRange();
            headerRange.ClearContents();
            headerRange.SetHeaderFormat();
            headerRange.Resize[1, 1].Value2 = BexConstants.IndividualLossSetName;

            var thresholdRangeName = GetThresholdRangeName(SegmentId, ComponentId);
            ReformatThreshold(thresholdRangeName);


            var inputLabelRange = GetInputLabelRange();
            inputLabelRange.SetBorderAroundToOrdinary();
            inputLabelRange.SetInputLabelColor();
            inputLabelRange.Locked = true;
            var labels = GetColumnLabels();
            inputLabelRange.Value = labels.ToOneByNArray();
            inputLabelRange.AlignRight();

            var inputRange = GetInputRange();
            var rangeWithoutTotal = inputRange.RemoveLastRow();
            SetInputInteriorColorContemplatingEstimate(rangeWithoutTotal);
            rangeWithoutTotal.SetInputFontColor();
            rangeWithoutTotal.SetBorderAroundToResizable();

            var labelColumnIndices = FindLabelColumnIndices();
            labelColumnIndices.Strings.ForEach(index =>
            {
                rangeWithoutTotal.GetColumn(index).NumberFormat = FormatExtensions.StringFormat;
                rangeWithoutTotal.GetColumn(index).AlignLeft();
                rangeWithoutTotal.GetColumn(index).ColumnWidth = ExcelConstants.StandardColumnWidth;
                inputLabelRange.GetColumn(index).AlignLeft();
            });
            labelColumnIndices.Dates.ForEach(index =>
            {
                rangeWithoutTotal.GetColumn(index).NumberFormat = FormatExtensions.DateFormat;
                rangeWithoutTotal.GetColumn(index).AlignRight();
                rangeWithoutTotal.GetColumn(index).ColumnWidth = ExcelConstants.StandardColumnWidth;
            });
            labelColumnIndices.Numbers.ForEach(index =>
            {
                rangeWithoutTotal.GetColumn(index).NumberFormat = FormatExtensions.WholeNumberFormat;
                rangeWithoutTotal.GetColumn(index).AlignRight();
                rangeWithoutTotal.GetColumn(index).ColumnWidth = ExcelConstants.StandardColumnWidth;
            });

            var rangeTotal = inputRange.GetLastRow();
            rangeTotal.ClearFontColor();
            rangeTotal.SetInputLabelInteriorColor();
            rangeTotal.ClearContents();
            rangeTotal.Locked = true;
            rangeTotal.ClearBorderAllButTop();

            var columnCountToTotal = GetColumnCountToTotal();
            var columnOffset = columnCountToTotal - 1;

            var sumRange = rangeTotal.GetTopRightCell().Offset[0, -columnOffset].Resize[1, columnCountToTotal];
            sumRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            sumRange.AlignRight();

            var sumContentRange = sumRange.Offset[-rangeWithoutTotal.Rows.Count, 0].Resize[rangeWithoutTotal.Rows.Count, 1];
            sumRange.Formula = $"=Sum({sumContentRange.Address[false, false]}";
            SetColumnWidth(inputRange);
        }

        private IList<string> GetColumnLabels()
        {
            var list = new List<string>();
            var desc = GetSegment().IndividualLossSetDescriptor;
            
            list.Add(BexConstants.OccurrenceIdName);
            list.Add(BexConstants.ClaimIdName);
            if (desc.IsEventCodeAvailable) list.Add(BexConstants.EventCodeName);
            list.Add(BexConstants.DescriptionName);
            if (desc.IsAccidentDateAvailable) list.Add(BexConstants.AccidentDateName);
            if (desc.IsPolicyDateAvailable) list.Add(BexConstants.PolicyDateName);
            if (desc.IsReportDateAvailable) list.Add(BexConstants.ReportDateName);
            if (desc.IsPolicyLimitAvailable) list.Add(BexConstants.LimitName);
            if (desc.IsPolicyAttachmentAvailable) list.Add(BexConstants.SirAttachmentName);

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
            var validation = new StringBuilder();

            ValidateThreshold(validation);

            Items = new List<IndividualLossModelPlus>();

            var inputRange = GetInputRange().RemoveLastRow();
            var rows = inputRange.GetContent();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var labelColumnIndices = FindLabelColumnIndices();
            var dateRows = rows.GetColumns(labelColumnIndices.Dates.ToArray());
            var numberRows = rows.GetColumns(labelColumnIndices.Numbers.ToArray());

            var dateRowsAsDates = dateRows.ForceContentToNullableDates();
            var numberRowsAsDoubles = numberRows.ForceContentToDoubles();

            var lossSet = (IndividualLossSet) GetParent();
            var individualLossLedger = lossSet.Ledger;
            
            var descriptor = GetSegment().IndividualLossSetDescriptor;
            var isPolicyLimitAvailable = descriptor.IsPolicyLimitAvailable;
            var isPolicyAttachmentAvailable = descriptor.IsPolicyAttachmentAvailable;
            var isPaidAvailable = descriptor.IsPaidAvailable;
            var isLossAndAlaeCombined = descriptor.IsLossAndAlaeCombined;
            var isLossAndAlaeSeparated = !descriptor.IsLossAndAlaeCombined;
            var isAccidentDateAvailable = descriptor.IsAccidentDateAvailable;
            var isPolicyDateAvailable = descriptor.IsPolicyDateAvailable;
            var isReportDateAvailable = descriptor.IsReportDateAvailable;
            var isEventCodeAvailable = descriptor.IsEventCodeAvailable;


            var relativeColumn = 0;
            var relativeColumnWithinDataType = 0;
            var columnMetrics = new Dictionary<ColumnMetric, Metric>
            {
                {ColumnMetric.OccurrenceId, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++)},
                {ColumnMetric.ClaimNumber, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++)}
            };
            if (isEventCodeAvailable) columnMetrics.Add(ColumnMetric.EventCode, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            columnMetrics.Add(ColumnMetric.LossDescription, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType));
            
            relativeColumnWithinDataType = 0;
            if (isAccidentDateAvailable) columnMetrics.Add(ColumnMetric.AccidentDate, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            if (isPolicyDateAvailable) columnMetrics.Add(ColumnMetric.PolicyDate, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            if (isReportDateAvailable) columnMetrics.Add(ColumnMetric.ReportedDate, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType));

            relativeColumnWithinDataType = 0;
            if (isPolicyLimitAvailable) columnMetrics.Add(ColumnMetric.PolicyLimit, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            if (isPolicyAttachmentAvailable) columnMetrics.Add(ColumnMetric.PolicyAttachment, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            if (isPaidAvailable)
            {
                if (isLossAndAlaeSeparated) columnMetrics.Add(ColumnMetric.PaidLoss, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
                if (isLossAndAlaeSeparated) columnMetrics.Add(ColumnMetric.PaidAlae, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
                if (isLossAndAlaeCombined) columnMetrics.Add(ColumnMetric.PaidLossAndAlae, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            }
            if (isLossAndAlaeSeparated) columnMetrics.Add(ColumnMetric.ReportedLoss, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            if (isLossAndAlaeSeparated) columnMetrics.Add(ColumnMetric.ReportedAlae, new Metric(startColumn, relativeColumn++, relativeColumnWithinDataType++));
            if (isLossAndAlaeCombined) columnMetrics.Add(ColumnMetric.ReportedLossAndAlae, new Metric(startColumn, relativeColumn, relativeColumnWithinDataType));


            var rowCount = rows.GetLength(0);
            if (!individualLossLedger.Any()) individualLossLedger.Create(rowCount);
            
            var lossSetId = lossSet.SourceId;
            for (var row = 0; row < rowCount; row++)
            {
                var rowNumber = row + startRow;

                var rowFromExcel = rows.GetRow(row);

                var dateRow = dateRows.GetRow(row);
                var dateRowAsDate = dateRowsAsDates.GetRow(row);

                #region non dates in date cells
                var accidentDate = isAccidentDateAvailable ? dateRow[columnMetrics[ColumnMetric.AccidentDate].RelativeColumnWithinDataType] : null;
                var policyDate = isPolicyDateAvailable ? dateRow[columnMetrics[ColumnMetric.PolicyDate].RelativeColumnWithinDataType] : null;
                var reportDate = isReportDateAvailable ? dateRow[columnMetrics[ColumnMetric.ReportedDate].RelativeColumnWithinDataType] : null;
                
                var accidentDateAsDate = isAccidentDateAvailable ? dateRowAsDate[columnMetrics[ColumnMetric.AccidentDate].RelativeColumnWithinDataType] : new DateTime?();
                var policyDateAsDate = isPolicyDateAvailable ? dateRowAsDate[columnMetrics[ColumnMetric.PolicyDate].RelativeColumnWithinDataType] : new DateTime?();
                var reportDateAsDate = isReportDateAvailable ? dateRowAsDate[columnMetrics[ColumnMetric.ReportedDate].RelativeColumnWithinDataType] : new DateTime?();

                var historicalPeriodType = GetSegment().HistoricalPeriodType;
                var dateMessageAlreadyLogged = false;

                if (accidentDate != null && accidentDateAsDate == null)
                {
                    var columnLetter = columnMetrics[ColumnMetric.AccidentDate].ColumnLetter;
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                    var message = $"{BexConstants.AccidentDateName.ToStartOfSentence()} <{accidentDate}> in {addressLocation} {BexConstants.NotRecognizedAsADate}";
                    validation.AppendLine(message);
                    if (historicalPeriodType.Equals("1")) dateMessageAlreadyLogged = true;
                }

                if (policyDate != null && policyDateAsDate == null)
                {
                    var columnLetter = columnMetrics[ColumnMetric.PolicyDate].ColumnLetter;
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                    var message = $"{BexConstants.PolicyDateName.ToStartOfSentence()} <{policyDate}> in {addressLocation} {BexConstants.NotRecognizedAsADate}";
                    validation.AppendLine(message);
                    if (historicalPeriodType.Equals("2")) dateMessageAlreadyLogged = true;
                }

                if (reportDate != null && reportDateAsDate == null)
                {
                    var columnLetter = columnMetrics[ColumnMetric.ReportedDate].ColumnLetter;
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                    var message = $"{BexConstants.ReportDateName.ToStartOfSentence()} <{reportDate}> in {addressLocation} {BexConstants.NotRecognizedAsADate}";
                    validation.AppendLine(message);
                    if (historicalPeriodType.Equals("3")) dateMessageAlreadyLogged = true;
                }

                #endregion

                #region non numbers in number cells

                var amountRow = numberRows.GetRow(row);
                var amountRowAsDoubles = numberRowsAsDoubles.GetRow(row);
                
                var limit = isPolicyLimitAvailable ? amountRow[columnMetrics[ColumnMetric.PolicyLimit].RelativeColumnWithinDataType] : null;
                var attachment = isPolicyAttachmentAvailable ? amountRow[columnMetrics[ColumnMetric.PolicyAttachment].RelativeColumnWithinDataType] : null;
                var paidLoss = isPaidAvailable && isLossAndAlaeSeparated ? amountRow[columnMetrics[ColumnMetric.PaidLoss].RelativeColumnWithinDataType] : null;
                var paidAlae = isPaidAvailable && isLossAndAlaeSeparated ? amountRow[columnMetrics[ColumnMetric.PaidAlae].RelativeColumnWithinDataType] : null;
                var paidLossAndAlae = isPaidAvailable && isLossAndAlaeCombined ? amountRow[columnMetrics[ColumnMetric.PaidLossAndAlae].RelativeColumnWithinDataType] : null;
                var reportedLoss = isLossAndAlaeSeparated ? amountRow[columnMetrics[ColumnMetric.ReportedLoss].RelativeColumnWithinDataType] : null;
                var reportedAlae = isLossAndAlaeSeparated ? amountRow[columnMetrics[ColumnMetric.ReportedAlae].RelativeColumnWithinDataType] : null;
                var reportedLossAndAlae = isLossAndAlaeCombined ? amountRow[columnMetrics[ColumnMetric.ReportedLossAndAlae].RelativeColumnWithinDataType] : null;
                
                var limitDouble = isPolicyLimitAvailable ? amountRowAsDoubles[columnMetrics[ColumnMetric.PolicyLimit].RelativeColumnWithinDataType] : double.NaN;
                var attachmentDouble = isPolicyAttachmentAvailable ? amountRowAsDoubles[columnMetrics[ColumnMetric.PolicyAttachment].RelativeColumnWithinDataType] : double.NaN;
                var paidLossAsDouble = isPaidAvailable && isLossAndAlaeSeparated ? amountRowAsDoubles[columnMetrics[ColumnMetric.PaidLoss].RelativeColumnWithinDataType] : double.NaN;
                var paidAlaeAsDouble = isPaidAvailable && isLossAndAlaeSeparated ? amountRowAsDoubles[columnMetrics[ColumnMetric.PaidAlae].RelativeColumnWithinDataType] : double.NaN;
                var paidLossAndAlaeAsDouble = isPaidAvailable && isLossAndAlaeCombined ? amountRowAsDoubles[columnMetrics[ColumnMetric.PaidLossAndAlae].RelativeColumnWithinDataType] : double.NaN;
                var reportedLossAsDouble = isLossAndAlaeSeparated ? amountRowAsDoubles[columnMetrics[ColumnMetric.ReportedLoss].RelativeColumnWithinDataType] : double.NaN;
                var reportedAlaeAsDouble = isLossAndAlaeSeparated ? amountRowAsDoubles[columnMetrics[ColumnMetric.ReportedAlae].RelativeColumnWithinDataType] : double.NaN;
                var reportedLossAndAlaeAsDouble = isLossAndAlaeCombined ? amountRowAsDoubles[columnMetrics[ColumnMetric.ReportedLossAndAlae].RelativeColumnWithinDataType] : double.NaN;
                
                if (limit != null && double.IsNaN(limitDouble))
                {
                    var columnLetter = columnMetrics[ColumnMetric.PolicyLimit].ColumnLetter;
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                    var message = $"{BexConstants.LimitName.ToStartOfSentence()} <{limit}> in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                }

                if (attachment != null && double.IsNaN(attachmentDouble))
                {
                    var columnLetter = columnMetrics[ColumnMetric.PolicyAttachment].ColumnLetter;
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                    var message = $"{BexConstants.SirAttachmentName.ToStartOfSentence()} <{attachment}> in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                }

                if (isLossAndAlaeCombined)
                {
                    if (isPaidAvailable)
                    {
                        if (paidLossAndAlae != null && double.IsNaN(paidLossAndAlaeAsDouble))
                        {
                            var columnLetter = columnMetrics[ColumnMetric.PaidLossAndAlae].ColumnLetter;
                            var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                            var message = $"{BexConstants.PaidLossAndAlaeName.ToStartOfSentence()} <{paidLossAndAlae}> in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                            validation.AppendLine(message);
                        }
                    }

                    if (reportedLossAndAlae != null && double.IsNaN(reportedLossAndAlaeAsDouble))
                    {
                        var columnLetter = columnMetrics[ColumnMetric.ReportedLossAndAlae].ColumnLetter;
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                        var message = $"{BexConstants.ReportedLossAndAlaeName.ToStartOfSentence()} <{reportedLossAndAlae}> in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                        validation.AppendLine(message);
                    }
                }
                else
                {
                    if (isPaidAvailable)
                    {
                        if (paidLoss != null && double.IsNaN(paidLossAsDouble))
                        {
                            var columnLetter = columnMetrics[ColumnMetric.PaidLoss].ColumnLetter;
                            var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                            var message = $"{BexConstants.PaidLossName.ToStartOfSentence()} <{paidLoss}> " +
                                          $"in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                            validation.AppendLine(message);
                        }

                        if (paidAlae != null && double.IsNaN(paidAlaeAsDouble))
                        {
                            var columnLetter = columnMetrics[ColumnMetric.PaidAlae].ColumnLetter;
                            var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                            var message = $"{BexConstants.PaidAlaeName.ToStartOfSentence()} <{paidAlae}> " +
                                          $"in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                            validation.AppendLine(message);
                        }
                    }

                    if (reportedLoss != null && double.IsNaN(reportedLossAsDouble))
                    {
                        var columnLetter = columnMetrics[ColumnMetric.ReportedLoss].ColumnLetter;
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                        var message = $"{BexConstants.ReportedLossName.ToStartOfSentence()} <{reportedLoss}> " +
                                      $"in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                        validation.AppendLine(message);
                    }

                    if (reportedAlae != null && double.IsNaN(reportedAlaeAsDouble))
                    {
                        var columnLetter = columnMetrics[ColumnMetric.ReportedAlae].ColumnLetter;
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                        var message = $"{BexConstants.ReportedAlaeName.ToStartOfSentence()} <{reportedAlae}> " +
                                      $"in {addressLocation} {BexConstants.NotRecognizedAsANumber}";
                        validation.AppendLine(message);
                    }
                }

                #endregion

                var occurrenceId = rowFromExcel[columnMetrics[ColumnMetric.OccurrenceId].RelativeColumnWithinDataType]?.ToString();
                var claimNumber = rowFromExcel[columnMetrics[ColumnMetric.ClaimNumber].RelativeColumnWithinDataType]?.ToString();
                var eventCode = isEventCodeAvailable ? rowFromExcel[columnMetrics[ColumnMetric.EventCode].RelativeColumnWithinDataType]?.ToString() : string.Empty;
                var lossDescription = rowFromExcel[columnMetrics[ColumnMetric.LossDescription].RelativeColumnWithinDataType]?.ToString();
                
                var individualLoss = new IndividualLossModelPlus
                {
                    RowId = row,
                    RowNumber = rowNumber,
                    SourceId = individualLossLedger[row].SourceId,
                    IsDirty = individualLossLedger[row].IsDirty,

                    DataSetId = lossSetId,
                    
                    OccurrenceId = occurrenceId,
                    ClaimNumber = claimNumber,
                    EventCode = eventCode,
                    LossDescription = lossDescription,

                    AccidentDate = accidentDateAsDate.MapToOffset(),
                    PolicyDate = policyDateAsDate.MapToOffset(),
                    ReportedDate = reportDateAsDate.MapToOffset(),
                    EvaluationDate = DateTime.Now.MapToOffset(),

                    PolicyLimitAmount = limit != null ? limitDouble : new double?(),
                    PolicyAttachmentAmount = attachment != null ? attachmentDouble : new double?(),

                    PaidLossAmount = paidLoss != null ? paidLossAsDouble : new double?(),
                    PaidAlaeAmount = paidAlae != null ? paidAlaeAsDouble : new double?(),
                    PaidCombinedAmount = paidLossAndAlae != null ? paidLossAndAlaeAsDouble : new double?(),

                    ReportedLossAmount = reportedLoss != null ? reportedLossAsDouble : new double?(),
                    ReportedAlaeAmount = reportedAlae != null ? reportedAlaeAsDouble : new double?(),
                    ReportedCombinedAmount = reportedLossAndAlae != null ? reportedLossAndAlaeAsDouble : new double?(),
                };

                //use to enforce at least one valid loss value - now allowing all/any loss values to be null
                var isDateValid = individualLoss.IsDateValid(historicalPeriodType);
                var isAnyContent = individualLoss.IsAnyContent();
                var historicalDateName = HistoricalPeriodTypesFromBex.GetDateFieldName(Convert.ToInt32(GetSegment().HistoricalPeriodType));

                if (!isAnyContent) continue;
                
                if (!isDateValid && !dateMessageAlreadyLogged)
                {
                    validation.AppendLine($"{historicalDateName.ToStartOfSentence()} in row {rowNumber} can't be blank");
                    continue;
                }

                Items.Add(individualLoss);
            }

            return validation;
        }

        public IModel GetParent()
        {
            return GetSegment().IndividualLossSets.Single(x => x.ComponentId == ComponentId);
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }

        public override Range GetBodyRange()
        {
            return GetInputRange();
        }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maximumDisplayOrder = GetSegment().IndividualLossSets.Max(x => x.IntraDisplayOrder);

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

                MessageHelper.Show($"Can't move the left-most {FriendlyName.ToLower()} farther to the left.", MessageType.Stop);
                return false;
            }
        }

        public override void MoveRight()
        {
            var segment = GetSegment();
            var siblingExcelMatrices = segment.IndividualLossSets.Select(a => a.ExcelMatrix).ToList();

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
            var siblingExcelMatrices = segment.IndividualLossSets.Select(a => a.ExcelMatrix).ToList();

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

        public int GetColumnCountToTotal()
        {
            var columnCount = 1;
            var individualLossDescriptor = GetSegment().IndividualLossSetDescriptor;

            if (individualLossDescriptor.IsPaidAvailable) columnCount *= 2;
            if (!individualLossDescriptor.IsLossAndAlaeCombined) columnCount *= 2;

            return columnCount;
        }

        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().IndividualLossSets.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1)
                .ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().IndividualLossSets.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1)
                .ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }

        public void ReformatBorderTop()
        {
            RangeName.GetRange().GetFirstRow().SetBorderTopToOrdinary();
        }

        public void ReformatBorderBottom()
        {
            RangeName.GetRange().GetLastRow().SetBorderBottomToOrdinary();
        }

        public static string GetRangeHeaderName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.IndividualLossSetRangeName}{componentId}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return GetHeaderRangeName(segmentId, componentId, ExcelConstants.IndividualLossSetRangeName);
        }

        public static string GetBasisRangeName(int segmentId, int componentId)
        {
            return $"segment{segmentId}.{ExcelConstants.IndividualLossSetRangeName}{componentId}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return $"{GetSublinesRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetThresholdRangeName(int segmentId, int componentId)
        {
            return $"segment{segmentId}.{ExcelConstants.IndividualLossSetRangeName}{componentId}.{ExcelConstants.ThresholdRangeName}";
        }

        public static int GetComponentIdFromThresholdRangeName(string thresholdRangeName)
        {
            var array = thresholdRangeName.Split('.');
            var componentId = Convert.ToInt32(array[1].Replace(ExcelConstants.IndividualLossSetRangeName, string.Empty));

            return componentId;
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.SublinesRangeName}";
        }

        public void ModifyRangeToReflectChangeToLimit(bool isLimitAvailable)
        {
            if (isLimitAvailable)
            {
                InsertLimitColumn();
            }
            else
            {
                DeleteColumn(BexConstants.LimitName);
            }
        }

        public void ModifyRangeToReflectChangeToAttachment(bool isAttachmentAvailable)
        {
            if (isAttachmentAvailable)
            {
                InsertAttachmentColumn();
            }
            else
            {
                DeleteColumn(BexConstants.SirAttachmentName);
            }
        }

        public void ModifyRangeToReflectChangeToAccidentDate(bool isAccidentDateAvailable)
        {
            if (isAccidentDateAvailable)
            {
                InsertAccidentDateColumn();
            }
            else
            {
                DeleteColumn(BexConstants.AccidentDateName);
            }
        }

        public void ModifyRangeToReflectChangeToPolicyDate(bool isPolicyDateAvailable)
        {
            if (isPolicyDateAvailable)
            {
                InsertPolicyDateColumn();
            }
            else
            {
                DeleteColumn(BexConstants.PolicyDateName);
            }
        }

        public void ModifyRangeToReflectChangeToReportDate(bool isReportDateAvailable)
        {
            if (isReportDateAvailable)
            {
                InsertReportDateColumn();
            }
            else
            {
                DeleteColumn(BexConstants.ReportDateName);
            }
        }

        public void ModifyRangeToReflectChangeToEventCode(bool isEventCodeAvailable)
        {
            if (isEventCodeAvailable)
            {
                InsertEventCodeColumn();
            }
            else
            {
                DeleteColumn(BexConstants.EventCodeName);
            }
        }

        protected override void DeletePaidColumns()
        {
            if (!DoLabelsContainAnExactMatch(new List<string> { BexConstants.PaidLossName, BexConstants.PaidLossAndAlaeName })) return; 
            
            var labelRange = GetInputLabelRange();
            var columnIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.PaidName).ToList();

            columnIndices.Reverse();
            foreach (var columnIndex in columnIndices)
            {
                labelRange.Resize[1, 1].Offset[0, columnIndex].EntireColumn.Delete();
            }
        }

        protected override void InsertPaidColumns()
        {
            if (DoLabelsContainAnExactMatch(new List<string> { BexConstants.PaidLossName, BexConstants.PaidLossAndAlaeName })) return; 
            
            var labelRange = GetInputLabelRange();
            var reportedColumnIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.ReportedName).ToList();
            var additionalColumnCount = reportedColumnIndices.Count;

            var firstReportedColumnIndex = reportedColumnIndices.First();
            labelRange.Offset[0, firstReportedColumnIndex].Resize[1, additionalColumnCount].InsertColumnsToRight();

            labelRange = GetInputLabelRange();
            for (var counter = 0; counter < additionalColumnCount; counter++)
            {
                var paidColumnOffset = additionalColumnCount + counter;
                var reportedColumnOffset = counter;
                var topRightRange = labelRange.GetTopRightCell();
                var paidLabel = topRightRange.Offset[0, -reportedColumnOffset].Value.ToString()
                    .Replace(BexConstants.ReportedName, BexConstants.PaidName);
                topRightRange.Offset[0, -paidColumnOffset].Value = paidLabel;
                topRightRange.Offset[0, -paidColumnOffset].ColumnWidth = topRightRange.ColumnWidth;
            }
        }

        public override bool HasData
        {
            get
            {
                var content = GetInputRange().RemoveLastRow().GetContent();
                for (var row = 0; row < content.GetLength(0); row++)
                {
                    for (var column = 0; column < content.GetLength(1); column++)
                    {
                        if (content[row, column] != null) return true;
                    }
                }
                return false;
            }
        }

        private void ReformatThreshold(string thresholdRangeName)
        {
            if (!thresholdRangeName.ExistsInWorkbook()) return;

            var thresholdRange = thresholdRangeName.GetRange();
            thresholdRange.SetBorderAroundToOrdinary();
            thresholdRange.Font.Bold = false;

            var labelRange = thresholdRange.GetTopLeftCell();
            labelRange.SetInputLabelFormatWithBorder();
            labelRange.Value2 = BexConstants.ThresholdName;
            labelRange.AlignLeft();
            labelRange.ClearBorderRight();
            
            var valueRange = thresholdRange.GetTopRightCell();
            valueRange.SetInputFormat();
            valueRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            valueRange.AlignRight();
            valueRange.ClearBorderLeft();
            SetInputInteriorColorContemplatingEstimate(valueRange);
        }

        private void ValidateThreshold(StringBuilder validation)
        {
            var thresholdRangeName = GetThisThresholdRangeName();
            if (!thresholdRangeName.ExistsInWorkbook()) return;

            var thresholdRange = thresholdRangeName.GetRange();
            var value = thresholdRange.GetTopRightCell().Value2;
            if (value == null)
            {
                Threshold = new int?(); 
                return;
            }

            if (double.TryParse(value.ToString(), out double theResult))
            {
                Threshold = Convert.ToInt32(Math.Round(theResult, 0, MidpointRounding.AwayFromZero));
            }
            else
            {
                validation.AppendLine($"{BexConstants.ThresholdName} <{value}> is not a number");
            }
        }

        private string GetThisThresholdRangeName()
        {
            return GetThresholdRangeName(SegmentId, ComponentId);
        }

        private enum ColumnMetric
        {
            OccurrenceId,
            ClaimNumber,
            EventCode,
            LossDescription,
            AccidentDate,
            PolicyDate,
            ReportedDate,
            PolicyLimit,
            PolicyAttachment,
            ReportedLoss,
            ReportedAlae,
            ReportedLossAndAlae,
            PaidLoss,
            PaidAlae,
            PaidLossAndAlae
        }

        private class Metric
        {
            public Metric(int startColumn, int relativeColumn, int relativeColumnWithinDataType)
            {
                RelativeColumn = relativeColumn;
                RelativeColumnWithinDataType = relativeColumnWithinDataType;

                var absoluteColumn = RelativeColumn + startColumn;
                ColumnLetter = absoluteColumn.GetColumnLetter();
            }

            private int RelativeColumn { get; }
            internal int RelativeColumnWithinDataType { get; }
            internal string ColumnLetter { get; }

        }

        private void SetColumnWidth(Range range)
        {
            range.ColumnWidth = ExcelConstants.StandardColumnWidth;
            var lossIndices = FindLabelColumnWithLoss();
            lossIndices.ForEach(index =>
            {
                range.GetColumn(index).ColumnWidth = ExcelConstants.LossColumnWidth;
            });
            range.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;
        }
        
        private void InsertLimitColumn()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.LimitName)) return;

            var maximumDateIndex = FindMaximumLabelIndicesMatchPartial(BexConstants.DateName);
            var index = maximumDateIndex + 1;

            GetInputLabelRange().Offset[0, index].Resize[1, 1].InsertColumnsToRight();
            GetInputLabelRange().Resize[1, 1].Offset[0, index].Value2 = BexConstants.LimitName;

            GetInputRange().GetColumn(index).NumberFormat = FormatExtensions.WholeNumberFormat;
        }

        private void InsertAttachmentColumn()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.SirAttachmentName)) return;

            var maximumDateIndex = FindMaximumLabelIndicesMatchPartial(BexConstants.DateName);
            var index = maximumDateIndex + 1;
            if (GetSegment().IndividualLossSetDescriptor.IsPolicyLimitAvailable) index++;

            GetInputLabelRange().Offset[0, index].Resize[1, 1].InsertColumnsToRight();
            GetInputLabelRange().Resize[1, 1].Offset[0, index].Value2 = BexConstants.SirAttachmentName;

            GetInputRange().GetColumn(index).NumberFormat = FormatExtensions.WholeNumberFormat;
        }

        
        private void InsertAccidentDateColumn()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.AccidentDateName)) return; 
            
            var maximumDateIndex = FindMinimumLabelIndicesMatchPartial(BexConstants.DateName);
            var index = maximumDateIndex;

            GetInputLabelRange().Offset[0, index].Resize[1, 1].InsertColumnsToRight();
            GetInputLabelRange().Resize[1, 1].Offset[0, index].Value2 = BexConstants.AccidentDateName;
            GetInputLabelRange().Resize[1, 1].Offset[0, index].AlignRight();

            GetInputRange().GetColumn(index).AppendColumn().NumberFormat = FormatExtensions.DateFormat;
        }
        
        private void InsertPolicyDateColumn()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.PolicyDateName)) return; 
            
            var dateLabelIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.DateName).ToList();
            int index;
            if (dateLabelIndices.Count == 2)
            {
                index = dateLabelIndices[1];
            }
            else
            {
                var segment = GetSegment();
                index = segment.IndividualLossSetDescriptor.IsAccidentDateAvailable ? dateLabelIndices[0] + 1 : dateLabelIndices[0];
            }

            GetInputLabelRange().Offset[0, index].Resize[1, 1].InsertColumnsToRight();
            GetInputLabelRange().Resize[1, 1].Offset[0, index].Value2 = BexConstants.PolicyDateName;
            GetInputLabelRange().Resize[1, 1].Offset[0, index].AlignRight();

            GetInputRange().GetColumn(index).AppendColumn().NumberFormat = FormatExtensions.DateFormat;
        }

        private void DeleteColumn(string columnLabel)
        {
            if (!DoLabelsContainAnExactMatch(columnLabel)) return;

            var index = FindLabelIndex(columnLabel);
            GetInputLabelRange().Resize[1, 1].Offset[0, index].EntireColumn.Delete();
        }

        private void InsertReportDateColumn()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.ReportDateName)) return; 
            
            var maximumDateIndex = FindMaximumLabelIndicesMatchPartial(BexConstants.DateName);
            var index = maximumDateIndex + 1;

            GetInputLabelRange().Offset[0, index].Resize[1, 1].InsertColumnsToRight();
            GetInputLabelRange().Resize[1, 1].Offset[0, index].Value2 = BexConstants.ReportDateName;
            GetInputLabelRange().Resize[1, 1].Offset[0, index].AlignRight();

            GetInputRange().GetColumn(index).AppendColumn().NumberFormat = FormatExtensions.DateFormat;
        }

        private void InsertEventCodeColumn()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.EventCodeName)) return;

            //can use index because labels to left are always present
            var index = Convert.ToInt32(ColumnMetric.EventCode);

            GetInputLabelRange().Offset[0, index].Resize[1, 1].InsertColumnsToRight();
            GetInputLabelRange().Resize[1, 1].Offset[0, index].Value2 = BexConstants.EventCodeName;
            GetInputLabelRange().Resize[1, 1].Offset[0, index].AlignLeft();

            GetInputRange().GetColumn(index).AppendColumn().NumberFormat = FormatExtensions.StringFormat;
            this.HeaderRangeName.GetRange().AppendColumnsToLeft(1).Name = this.HeaderRangeName; 
 

        }


    private LabelColumnIndices FindLabelColumnIndices()
        {
            var columnIndices = new LabelColumnIndices();
            var columnCount = GetInputLabelRange().Columns.Count;
            var indices = Enumerable.Range(0, columnCount);

            var dates = FindAllLabelIndicesWithPartialMatch(BexConstants.DateName).ToList();
            var numbers = FindAllLabelIndicesMatchPartial(NumberLabels).ToList();

            columnIndices.Dates = dates;
            columnIndices.Numbers = numbers;
            columnIndices.Strings = indices.Except(dates).Except(numbers);
            return columnIndices;
        }

        private List<int> FindLabelColumnWithLoss()
        {
            var indices = new List<int>();

            var labelsRange = GetInputLabelRange();
            var labels = labelsRange.GetContent().ForceContentToStrings().GetRow(0).ToList();

            var counter = 0;
            foreach (var label in labels)
            {
                if (label.Contains("Loss") || label.Contains("ALAE")) indices.Add(counter);
                counter++;
            }
            return indices;
        }

        public override void ToggleEstimate()
        {
            IsEstimate = !IsEstimate;

            var inputRange = GetInputRange().RemoveLastRow();
            SetInputInteriorColorContemplatingEstimate(inputRange);

            var thresholdRangeName = GetThisThresholdRangeName();
            if (!thresholdRangeName.ExistsInWorkbook()) return;

            var thresholdRange = thresholdRangeName.GetRange();
            SetInputInteriorColorContemplatingEstimate(thresholdRange.GetTopRightCell());
        }

        private struct LabelColumnIndices
        {
            internal IEnumerable<int> Dates { get; set; }
            internal IEnumerable<int> Numbers { get; set; }
            internal IEnumerable<int> Strings { get; set; }
        }

    }

}



