using SubmissionCollector.Enums;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal class PolicyProfileColumnDeleter : BaseColumnDeleter
    {
        protected override int MinimumInputColumnCount => 5;
        protected override int LabelColumnCount => 0;
        protected override int FrozenLeftColumnCount => 2;
        protected override int FrozenRightColumnCount => 2;

        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            if (((PolicyExcelMatrix) excelMatrix).Dimension is FlatPolicyProfileDimension)
            {
                var message = $"In order to delete columns, the selected {PionlearClient.BexConstants.PolicyProfileName.ToLower()} can't be Flat";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            return base.Validate(excelMatrix, range);
        }
    }
}
