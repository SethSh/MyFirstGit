using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;

namespace SubmissionCollector.Models.DataComponents
{
    public interface IProfileFormatter
    {
        bool RequiresNormalization { get;}
        void FormatDataRange(Range dataRange);
        void WriteSumFormulaToRange(Range sumRange, string sumFormula);
    }

    public abstract class BaseProfileFormatter : IProfileFormatter
    {
        public abstract bool RequiresNormalization { get; }

        public virtual void FormatDataRange(Range dataRange)
        {
            dataRange.Validation.Delete();
        }

        public virtual void WriteSumFormulaToRange(Range sumRange, string sumFormula)
        {
            sumRange.Formula = sumFormula;
            sumRange.ClearSumCheck();
            sumRange.AlignRight();
        }
    }

    public class PercentProfileFormatter : BaseProfileFormatter
    {
        public override void WriteSumFormulaToRange(Range sumRange, string sumFormula)
        {
            base.WriteSumFormulaToRange(sumRange, sumFormula);
            sumRange.ApplySumCheck();
            sumRange.NumberFormat = FormatExtensions.PercentFormat;
        }


        public override bool RequiresNormalization => false; 

        public override void FormatDataRange(Range dataRange)
        {
            dataRange.NumberFormat = FormatExtensions.PercentFormat;
            base.FormatDataRange(dataRange);
        }
    }

    internal class PremiumProfileFormatter : BaseProfileFormatter
    {
        public override bool RequiresNormalization => true; 

        public override void FormatDataRange(Range dataRange)
        {
            dataRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            base.FormatDataRange(dataRange);
        }

        public override void WriteSumFormulaToRange(Range sumRange, string sumFormula)
        {
            base.WriteSumFormulaToRange(sumRange, sumFormula);
            sumRange.NumberFormat = FormatExtensions.WholeNumberFormat;
        }
    }
}