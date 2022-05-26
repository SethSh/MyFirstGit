using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class PeriodFillManager : BaseFillManger
    {

        public PeriodFillManager(IExcelMatrix excelMatrix) : base (excelMatrix)
        {
        
        }

        public override void Fill(IWorkbookLogger logger, bool values)
        {
            try
            {
                var range = ExcelMatrix.GetInputRange();
                var rowCount = range.Rows.Count;

                var firstRowStartRange = range.GetTopLeftCell();
                var firstRowEndRange = firstRowStartRange.Offset[0, 1];
                var firstRowEvaluationRange = firstRowStartRange.Offset[0, 2];

                if (GetTopLeftContent(range) == null)
                {
                    MessageHelper.Show($"Enter the first {BexConstants.StartDateName.ToLower()} in range {firstRowStartRange.Address}", MessageType.Warning);
                    return;
                }

                if (GetTopLeftDate(range) == null)
                {
                    MessageHelper.Show($"Enter the first {BexConstants.StartDateName.ToLower()} with a valid date in range {firstRowStartRange.Address}", MessageType.Warning);
                    return;
                }

                var firstRowEvaluationDate = GetFirstEvaluationDate(range);
                var defaultEvaluationDate = GetDefaultEvaluationDate();

                if (firstRowEvaluationDate == null && defaultEvaluationDate == null)
                {
                    MessageHelper.Show($"Either enter the first {BexConstants.EvaluationDateName.ToLower()} in range {firstRowEvaluationRange.Address} " +
                                       $"or go to {BexConstants.UserPreferencesName} and select a {BexConstants.EvaluationDateName.ToLower()}", MessageType.Warning);
                    return;
                }

                if (IsPopulated(range.Offset[0, 1].Resize[rowCount - 1, 3]))
                {
                    const string message = "This button will overwrite all rows but the first.  Are you sure you want to continue?";
                    if (MessageHelper.ShowWithYesNo(message) != DialogResult.Yes) return;
                }

                using (new ExcelScreenUpdateDisabler())
                {
                    var address = firstRowStartRange.Address[false, false];
                    if (GetFirstEndDate(range) == null)
                    {
                        firstRowEndRange.Resize[rowCount, 1].Formula = $"=Date(Year({address})+1, Month({address}), Day({address})) - 1";
                    }
                    else
                    {
                        var address2 = firstRowStartRange.Offset[1, 0].Address[false, false];
                        firstRowEndRange.Offset[1, 0].Resize[rowCount - 1, 1].Formula = $"=Date(Year({address2})+1, Month({address2}), Day({address2})) - 1";
                    }

                    address = firstRowEndRange.Address[false, false];
                    firstRowStartRange.Offset[1, 0].Resize[rowCount - 1, 1].Formula = $"={address}+1";

                    if (firstRowEvaluationDate == null) firstRowEvaluationRange.Value2 = defaultEvaluationDate;
                    firstRowEvaluationRange.Offset[1, 0].Resize[range.Rows.Count - 1, 1].Formula = $"={firstRowEvaluationRange.Address[false, false]}";

                    if (!values) return;

                    firstRowStartRange.Offset[1, 0].Resize[rowCount - 1, 1].Value2 = firstRowStartRange.Offset[1, 0].Resize[rowCount - 1, 1].Value2;
                    firstRowEndRange.Resize[rowCount, 1].Value2 = firstRowEndRange.Resize[rowCount, 1].Value2;
                    firstRowEvaluationRange.Offset[1, 0].Resize[range.Rows.Count - 1, 1].Value2 = firstRowEvaluationRange.Offset[1, 0].Resize[range.Rows.Count - 1, 1].Value2;
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Quick fill dates failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }


        private static DateTime? GetFirstEndDate(Range inputRange)
        {
            var date = inputRange.GetTopLeftCell().Offset[0, 1];
            return date.GetContent().ForceContentToNullableDates()[0, 0];
        }

        private static DateTime? GetFirstEvaluationDate(Range inputRange)
        {
            var dateRange = inputRange.GetTopRightCell();
            var date = dateRange.GetContent().ForceContentToNullableDates();
            return date?[0, 0];
        }

        private static DateTime? GetDefaultEvaluationDate()
        {
            var evaluationDateType = UserPreferences.ReadFromFile().EvaluationDateType;
            DateTime? evaluationDate = null;
            switch (evaluationDateType)
            {
                case EvaluationDateType.None:
                    break;
                case EvaluationDateType.EndOfLastHalfYear:
                    evaluationDate = BaseUserPreferencesViewModel.GetLastDayOfLastHalfYear();
                    break;
                case EvaluationDateType.EndOfLastQuarter:
                    evaluationDate = BaseUserPreferencesViewModel.GetLastDayOfLastQuarter();
                    break;
                case EvaluationDateType.EndOfLastYear:
                    evaluationDate = BaseUserPreferencesViewModel.GetLastDayOfLastYear();
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            return evaluationDate;
        }
    }

    internal class RateChangeFillManager : BaseFillManger
    {
        public RateChangeFillManager(IExcelMatrix excelMatrix) : base (excelMatrix)
        {
               
        }

        public override void Fill(IWorkbookLogger logger, bool values)
        {
            try
            {
                var range = ExcelMatrix.GetInputRange().GetFirstColumn();
                var rowCount = range.Rows.Count;

                var firstRowRange = range.GetTopLeftCell();
                if (GetTopLeftContent(range) == null)
                {
                    MessageHelper.Show($"Enter the first {BexConstants.RateChangeName.ToLower()} date in range {firstRowRange.Address}", MessageType.Warning);
                    return;
                }

                if (GetTopLeftDate(range) == null)
                {
                    MessageHelper.Show($"Enter the first {BexConstants.RateChangeName.ToLower()} date with a valid date in range {firstRowRange.Address}", MessageType.Warning);
                    return;
                }

                if (IsPopulated(range.Offset[1, 0].Resize[rowCount - 1, 1]))
                {
                    const string message = "This button will overwrite all rows but the first.  Are you sure you want to continue?";
                    if (MessageHelper.ShowWithYesNo(message) != DialogResult.Yes) return;
                }

                using (new ExcelScreenUpdateDisabler())
                {
                    var address = firstRowRange.Address[false, false];
                    firstRowRange.Offset[1, 0].Resize[rowCount - 1, 1].Formula = $"=Date(Year({address})+1, Month({address}), Day({address}))";

                    if (!values) return;

                    firstRowRange.Offset[1, 0].Resize[rowCount - 1, 1].Value2 = firstRowRange.Offset[1, 0].Resize[rowCount - 1, 1].Value2;
                }

            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Quick fill dates failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }

    internal abstract class BaseFillManger : IFillManager
    {
        protected IExcelMatrix ExcelMatrix;

        protected BaseFillManger(IExcelMatrix excelMatrix)
        {
            ExcelMatrix = excelMatrix;
        }
        
        internal object GetTopLeftContent(Range range)
        {
            return range.GetTopLeftCell().GetContent()[0, 0];
        }

        internal DateTime? GetTopLeftDate(Range range)
        {
            var date = range.GetTopLeftCell();
            return date.GetContent().ForceContentToNullableDates()[0, 0];
        }

        internal bool IsPopulated(Range range)
        {
            var rangeContent = range.GetContent();
            for (var row = 0; row < rangeContent.GetLength(0); row++)
            {
                for (var column = 0; column < rangeContent.GetLength(1); column++)
                {
                    var item = rangeContent[row, column];
                    if (item != null) return true;
                }
            }
            return false;
        }

        public abstract void Fill(IWorkbookLogger logger, bool values);
    }

    internal interface IFillManager
    {
        void Fill(IWorkbookLogger logger, bool values);
    }

    internal interface IRangeFillable
    {

    }

    internal static class FillManagerFactory
    {
        internal static IFillManager Create(IWorkbookLogger logger)
        {
            try
            {
                var componentIdentifier = new SegmentExcelComponentIdentifier();
                if (!componentIdentifier.Validate()) return null;

                if (!(componentIdentifier.ExcelMatrix is IRangeFillable))
                {
                    MessageHelper.Show(@"The selection must be within a fill-able range", MessageType.Stop);
                    return null;
                } 
            
                switch (componentIdentifier.ExcelMatrix)
                {
                    case PeriodSetExcelMatrix _:
                        return new PeriodFillManager(componentIdentifier.ExcelMatrix);
                    case RateChangeExcelMatrix _:
                        return new RateChangeFillManager(componentIdentifier.ExcelMatrix);
                    default:
                        const string message = "Fill manager not found";
                        throw new ArgumentOutOfRangeException(message);
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Quick fill dates failed finding manager";
                MessageHelper.Show(message, MessageType.Stop);
                return null;
            }
        }
    }
}
