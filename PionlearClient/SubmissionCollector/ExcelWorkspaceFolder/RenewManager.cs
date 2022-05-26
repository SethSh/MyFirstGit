using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelUtilities.RangeSizeModifier;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class RenewManager
    {
        private const string Title = "Renew Validation Issues";
        private const string SuccessMessage = "Renew Successful";
        private const string FailureMessage = "Renew Failed";

        public static void Renew(IWorkbookLogger logger)
        {
            try
            {
                Renew();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show(FailureMessage, MessageType.Stop); 
            }
        }

        private static void Renew()
        {
                string message;
                if (WorkbookSaveManager.IsReadOnly)
                {
                    message = "Can't renew a read-only workbook";
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }

                var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
                var packageModel = package.CreatePackageModel();

                var validationResult = Validate(package, packageModel);
                if (!validationResult.IsValid) return;

                message = "Prior to renewing you must SAVE AS your workbook.  Have you already done the SAVE AS?";
                var dialogResult = MessageHelper.ShowWithYesNo(message);
                if (dialogResult != DialogResult.Yes) return;

                var segments = package.Segments;

                using (new ExcelEventDisabler())
                {
                    using (new ExcelScreenUpdateDisabler())
                    {
                        AdvanceUnderwritingYear(package, packageModel);

                        foreach (var segment in segments)
                        {
                            var rowCount = GetRowCount(segment);
                            var columnCount = GetColumnCount(segment);

                            if (segment.PeriodSet.ExcelMatrix.GetInputRange().Rows.Count == rowCount)
                            {
                                var lastRowContent = GetLastRowContent(segment, rowCount, columnCount);
                                DeleteLastRowContent(segment, rowCount, columnCount);

                                var rangeInserter = validationResult.RangerInserters[segment.Id];
                                rangeInserter.ModifyRange();

                                ReplaceLastRowContent(segment, rowCount, lastRowContent);
                            }

                            rowCount++;
                            FormatNewRow(segment, rowCount, columnCount);
                            AdvanceEvaluationDates(segment, rowCount);
                            AdvanceStartAndEndDates(segment, rowCount);
                        }
                    }
                }

                package.CopyIdsIntoPredecessorIds();
                if (package.SourceId.HasValue)
                {
                    package.DecoupleFromServer();
                    package.BexCommunications.Clear();
                    WorkbookSaveManager.Save();
                }

                MessageHelper.Show(SuccessMessage);
            
        }

        private static ValidationResult Validate(IPackage package, PackageModel packageModel)
        {
            if (package.ExcelValidation.Length > 0)
            {
                MessageHelper.Show(Title, package.ExcelValidation.ToString(), MessageType.Warning);
                return new ValidationResult {IsValid = false};
            }

            var validation = packageModel.Validate();
            if (validation.Length > 0)
            {
                MessageHelper.Show(Title, validation.ToString(), MessageType.Warning);
                return new ValidationResult { IsValid = false };
            }

            var inserterValidation = new StringBuilder();
            var rangerInserters = new Dictionary<int, PeriodSetRowInserter>();

            foreach (var segment in package.Segments)
            {
                var excelMatrix = segment.PeriodSet.ExcelMatrix;
                var range = excelMatrix.GetInputRange().GetLastRow();
                var rangeInserter = new PeriodSetRowInserter();

                if (!rangeInserter.Validate(excelMatrix, range))
                {
                    inserterValidation.AppendLine($"Segment {segment.Name}: can't append an additional {BexConstants.PeriodName.ToLower()}");
                }

                rangerInserters.Add(segment.Id, rangeInserter);
            }

            if (inserterValidation.Length == 0)
            {
                return  new ValidationResult {RangerInserters = rangerInserters, IsValid = true};
            }


            MessageHelper.Show(Title, inserterValidation.ToString(), MessageType.Warning);
            return new ValidationResult { IsValid = false };
        }

        private static void AdvanceUnderwritingYear(IPackage package, PackageModel packageModel)
        {
            var range = package.UnderwritingYearExcelMatrix.GetInputRange();
            range.Value2 = Convert.ToInt32(packageModel.UnderwritingYear) + 1;
            range.SetRenewedInputColor();
        }

        private static void AdvanceEvaluationDates(ISegment segment, int rowCount)
        {
            var evaluationDatesRange = segment.PeriodSet.ExcelMatrix.GetInputRange().GetLastColumn().Resize[rowCount, 1];
            var evaluationDates = evaluationDatesRange.GetContent().ForceContentToNullableDates();

            var newEvaluationDates = new object[rowCount, 1];

            var rowIndex = 0;
            foreach (var date in evaluationDates)
            {
                if (date.HasValue)
                {
                    newEvaluationDates[rowIndex, 0] = date.Value.AddYears(1);
                }

                rowIndex++;
            }

            var range = evaluationDatesRange;
            range.Value2 = newEvaluationDates;
            range.SetRenewedInputColor();
        }

        private static void ReplaceLastRowContent(ISegment segment, int rowCount, IList<string> lastRow)
        {
            //write in former last row into freshly inserted penultimate last row
            var range = segment.PeriodSet.ExcelMatrix.GetInputRange().GetRow(rowCount)
                .GetTopLeftCell().Offset[-1, 0].Resize[1, lastRow.Count];

            for (var column = 0; column < range.Columns.Count; column++)
            {
                range.GetTopLeftCell().Offset[0, column].Formula = lastRow[column];
            }
        }

        private static void FormatNewRow(ISegment segment, int rowCount, int columnCount)
        {
            var range = segment.PeriodSet.ExcelMatrix.GetInputRange().GetRow(rowCount - 1)
                .GetTopLeftCell().Offset[-1, 0].Resize[1, columnCount];

            range.Offset[1, 0].SetRenewedInputColor();
        }

        private static void DeleteLastRowContent(ISegment segment, int rowCount, int columnCount)
        {
            var range = segment.PeriodSet.ExcelMatrix.GetInputRange()
                .GetRow(rowCount-1).GetTopLeftCell().Resize[1, columnCount];

            range.ClearContents();
        }

        private static void AdvanceStartAndEndDates(ISegment segment, int rowCount)
        {
            var datesFromRange = segment.PeriodSet.ExcelMatrix.GetInputRange().GetRow(rowCount-1).Offset[-1, 0];
            var fromStartRange = datesFromRange.GetTopLeftCell();
            var fromEndRange = datesFromRange.GetRangeSubset(0, 1).GetTopLeftCell();

            var datesToRange = segment.PeriodSet.ExcelMatrix.GetInputRange().GetRow(rowCount - 1);
            var toStartRange = datesToRange.GetTopLeftCell();
            var toEndRange = datesToRange.GetRangeSubset(0, 1).GetTopLeftCell();
            var toEvaluationRange = datesToRange.GetRangeSubset(0, 2).GetTopLeftCell();

            var isStartDateFormula = fromStartRange.Value2.ToString() != fromStartRange.Formula.ToString();
            var isEndDateFormula = fromEndRange.Value2.ToString() != fromEndRange.Formula.ToString();

            var datesFrom = datesFromRange.GetContent().ForceContentToNullableDates();
            var startDateFrom = datesFrom[0, 0];
            var endDateFrom = datesFrom[0, 1];
            var evaluationDateFrom = datesFrom[0, 2];

            if (startDateFrom.HasValue)
            {
                if (isStartDateFormula)
                {
                    fromStartRange.Resize[2, 1].Formula = fromStartRange.Formula;
                }
                else
                {
                    var newStartDate = startDateFrom.Value.AddYears(1);
                    toStartRange.Value2 = newStartDate;
                }
            }

            if (endDateFrom.HasValue)
            {
                if (isEndDateFormula)
                {
                    fromEndRange.Resize[2, 1].Formula = fromEndRange.Formula;
                }
                else
                {
                    var newEndDate = endDateFrom.Value.AddYears(1);
                    toEndRange.Value2 = newEndDate;
                }
            }

            if (evaluationDateFrom.HasValue)
            {
                var newEvaluationDate = evaluationDateFrom.Value;
                toEvaluationRange.Value2 = newEvaluationDate;
            }

            datesToRange.SetRenewedInputColor();
        }

        private static IList<string> GetLastRowContent(ISegment segment, int rowCount, int columnCount)
        {
            var periodSetRange = segment.PeriodSet.ExcelMatrix.GetInputRange();
            var lastRowRange = periodSetRange.GetRow(rowCount-1).Resize[1, columnCount];
            var lastRowContent = lastRowRange.GetContentFormulas().GetRow(0);

            var content = lastRowContent.Select(item => item != null ? item.ToString() : string.Empty).ToList();
            return content;
        }

        private static int GetRowCount(ISegment segment)
        {
            var periodSetRange = segment.PeriodSet.ExcelMatrix.GetInputRange();
            var periodCornerRange = periodSetRange.GetTopLeftCell();

            var totalRowCount = periodSetRange.Rows.Count;
            var rowCount = totalRowCount;
            for (var row = 0; row <rowCount; row++)
            {
                if (periodCornerRange.Offset[totalRowCount - row - 1, 0].Value2 == null)
                {
                    rowCount -= 1;
                }
                else
                {
                    break;
                }
            }

            return rowCount;
        }
        private static int GetColumnCount(ISegment segment)
        {
            var periodSetRange = segment.PeriodSet.ExcelMatrix.GetInputRange();
            var columnCount = periodSetRange.Columns.Count
                              + segment.ExposureSets.Sum(set => set.ExcelMatrix.GetInputRange().Columns.Count)
                              + segment.AggregateLossSets.Sum(set => set.ExcelMatrix.GetInputRange().Columns
                                  .Count);


            return columnCount;
        }

        private struct ValidationResult
        {
            internal bool IsValid;
            internal IDictionary<int, PeriodSetRowInserter> RangerInserters;
        }
    }
}
