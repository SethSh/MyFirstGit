using System.Text;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;

namespace SubmissionCollector.Models.Package.DataComponents
{
    public class UnderwritingYearExcelMatrix : BasePackageExcelMatrix
    {
        public override string FriendlyName => BexConstants.UnderwritingYearName;
        public override string ExcelRangeName => ExcelConstants.UnderwritingYearRangeName;

        public override void Reformat()
        {
            base.Reformat();

            var range = RangeName.GetRange();
            range.ClearBorders();
            range.SetBorderToOrdinary();

            var labelRange = GetInputLabelRange();
            labelRange.Locked = true;
            labelRange.SetInputLabelColor();
            
            var inputRange = GetInputRange();
            inputRange.NumberFormat = FormatExtensions.YearFormat;
        }

        public override Range GetInputRange()
        {
            return RangeName.GetRange().GetTopRightCell();
        }

        public override Range GetInputLabelRange()
        {
            return RangeName.GetRange().GetTopLeftCell();
        }

        public override StringBuilder Validate()
        {
            throw new System.NotImplementedException();
        }
        
    }
}
