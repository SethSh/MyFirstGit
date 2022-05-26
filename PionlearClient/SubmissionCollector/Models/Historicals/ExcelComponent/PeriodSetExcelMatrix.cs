using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;
using Newtonsoft.Json;
using PionlearClient;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelWorkspaceFolder;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Historicals.ExcelComponent
{
    public class PeriodSetExcelMatrix : SingleOccurrenceSegmentExcelMatrix, IRangeResizable, IRangeFillable
    {
        public PeriodSetExcelMatrix(int segmentId) : base(segmentId)
        {

        }

        
        public override string FriendlyName => BexConstants.PeriodSetName;
        public override string ExcelRangeName => ExcelConstants.PeriodSetRangeName;

        [JsonIgnore]
        public Dictionary<int, ValidationDetail> ValidationDetails { get; set; }

        [JsonIgnore]
        public IList<CollectorApi.SubmissionSegmentPeriod> Items { set; get; }

        public override void Reformat()
        {
            base.Reformat();

            var range = GetInputRange();
            range.AlignRight();
            range.SetBorderAroundToResizable();
            range.SetBorderRightToOrdinary();
            range.NumberFormat = FormatExtensions.DateFormat;

            range.GetLastRow().Offset[1,0].ClearBorderAllButTop();
            
            range.ColumnWidth = ExcelConstants.StandardColumnWidth;
        }
        
        public void ReformatBorderTop()
        {
            GetInputRange().GetFirstRow().SetBorderTopToOrdinary();
        }

        public void ReformatBorderBottom()
        {
            GetInputRange().GetLastRow().SetBorderBottomToResizable();
        }
        
        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(1,0);
        }

        public override Range GetInputLabelRange()
        {
            throw new NotImplementedException();
        }

        public override StringBuilder Validate()
        {
            var validation = new StringBuilder();
            Items = new List<CollectorApi.SubmissionSegmentPeriod>();
            ValidationDetails = new Dictionary<int, ValidationDetail>();

            var inputRange = GetInputRange();

            var datesFromExcel = inputRange.GetContent();
            var dates = datesFromExcel.ForceContentToNullableDates();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            const int startDateColumn = 0;
            const int endDateColumn = 1;
            const int evalDateColumn = 2;

            var columnLetters = new Dictionary<int, string>
            {
                {startDateColumn, (startColumn + startDateColumn).GetColumnLetter()},
                {endDateColumn, (startColumn + endDateColumn).GetColumnLetter()},
                {evalDateColumn, (startColumn + evalDateColumn).GetColumnLetter()}
            };

            var rowCount = datesFromExcel.GetLength(0);

            
            for (var row = 0; row < rowCount; row++)
            {
                var startFromExcel = datesFromExcel[row, startDateColumn]?.ToString();
                var endFromExcel = datesFromExcel[row, endDateColumn]?.ToString();
                var evalFromExcel = datesFromExcel[row, evalDateColumn]?.ToString();

                if (startFromExcel == null && endFromExcel == null && evalFromExcel == null)
                {
                    ValidationDetails.Add(row, new ValidationDetail
                    {
                        IsNull = true,
                        StartDateColumnLetter = columnLetters[startDateColumn],
                        EndDateColumnLetter = columnLetters[endDateColumn],
                        EvaluationDateColumnLetter = columnLetters[evalDateColumn]
                    });
                    continue;
                }

                var start = dates[row, startDateColumn];
                var end = dates[row, endDateColumn];
                var eval = dates[row, evalDateColumn];

                if (start != null && end != null && eval != null)
                {
                    var item = new CollectorApi.SubmissionSegmentPeriod
                    {
                        StartDate = start.Value.MapToOffset(),
                        EndDate = end.Value.MapToOffset(),
                        EvaluationDate = eval.Value.MapToOffset()
                    };
                    Items.Add(item);

                    ValidationDetails.Add(row, new ValidationDetail
                    {
                        IsNull = false,
                        IsOk = true,
                        StartDate = start.Value,
                        EndDate = end.Value,
                        EvaluationDate = eval.Value,
                        StartDateColumnLetter = columnLetters[startDateColumn],
                        EndDateColumnLetter = columnLetters[endDateColumn],
                        EvaluationDateColumnLetter = columnLetters[evalDateColumn]
                    });
                }
                else
                {
                    var absoluteRow = row + startRow;
                    
                    if (start == null)
                    {
                        var columnLetter = columnLetters[startDateColumn];
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter,absoluteRow);
                        validation.AppendLine(startFromExcel == null
                            ? $"Enter effective date in {addressLocation}"
                            : $"Enter effective date in {addressLocation}: <{startFromExcel}> is not recognized as a date");
                    }
                    if (end == null)
                    {
                        var columnLetter = columnLetters[endDateColumn];
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);
                        validation.AppendLine(endFromExcel == null
                            ? $"Enter expiration date in {addressLocation}"
                            : $"Enter expiration date in {addressLocation}: <{endFromExcel}> is not recognized as a date");
                    }
                    if (eval == null)
                    {
                        var columnLetter = columnLetters[evalDateColumn];
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);
                        validation.AppendLine(evalFromExcel == null
                            ? $"Enter evaluation date in {addressLocation}"
                            : $"Enter evaluation date in {addressLocation}: <{evalFromExcel}> is not recognized as a date");
                    }
                    
                    ValidationDetails.Add(row, new ValidationDetail
                    {
                        IsNull = false,
                        IsOk = false,
                        StartDateColumnLetter = columnLetters[startDateColumn],
                        EndDateColumnLetter = columnLetters[endDateColumn],
                        EvaluationDateColumnLetter = columnLetters[evalDateColumn]
                    });
                }
            }

            return validation;
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }

        public override Range GetBodyRange()
        {
            return GetInputRange();
        }
    }

    public class ValidationDetail
    {
        public bool IsNull { get; set; }
        public bool IsOk { get; set; }

        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime EvaluationDate { get; set; }

        public string StartDateColumnLetter { get; set; }
        public string EndDateColumnLetter { get; set; }
        public string EvaluationDateColumnLetter { get; set; }

        public IList<string> ColumnLetters => new List<string> {StartDateColumnLetter, EndDateColumnLetter, EvaluationDateColumnLetter};
    }
}
