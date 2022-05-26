using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter
{
    public class SirByLimitPolicyProfileDimension : BasePolicyProfileDimension
    {
        //          limits      limit 1     limit 2
        //sirs                  1m          5m
        //sir 1     1000        30%       20%
        //sir 2     2000        15%       35%
        public SirByLimitPolicyProfileDimension(PolicyExcelMatrix policyExcelMatrix) : base(policyExcelMatrix)
        {

        }

        public override Range WeightRange => PolicyExcelMatrix.GetInputRange().GetRangeSubset(1,1);
        
        public override void Reformat()
        {
            PolicyExcelMatrix.GetSublinesRange().BorderAround2();
            PolicyExcelMatrix.GetSublinesHeaderRange().BorderAround2();
            PolicyExcelMatrix.GetBodyRange().Offset[-1, 0].BorderAround2();

            var bodyRange = PolicyExcelMatrix.GetBodyRange();
            bodyRange.ClearBorders();

            const int bufferRowCount = 1;
            var sirCount = bodyRange.Rows.Count - 1 - bufferRowCount;
            var limitCount = bodyRange.Columns.Count - 1;

            var header = PolicyExcelMatrix.GetBodyHeaderRange();
            var limits = bodyRange.Offset[0, 1].Resize[1, limitCount];
            var sirs = bodyRange.Offset[1, 0].Resize[sirCount, 1];
            var weights = bodyRange.Offset[1, 1].Resize[sirCount, limitCount];
            var buffer = bodyRange.GetLastRow();

            header.SetInputLabelInteriorColor();
            header.Value2 = $"SIR {RangeExtensions.DownArrow} Lim {RangeExtensions.RightArrow}";
            header.Locked = true;

            sirs.NumberFormat = FormatExtensions.WholeNumberFormat;
            sirs.SetBorderAroundToOrdinary();
            sirs.SetBorderBottomToResizable();
            
            limits.NumberFormat = FormatExtensions.WholeNumberFormat;
            limits.SetBorderAroundToOrdinary();
            limits.SetBorderRightToResizable();
            
            weights.NumberFormat = FormatExtensions.PercentFormat;
            weights.SetBorderRightToResizable();
            weights.SetBorderBottomToResizable();
            
            buffer.Locked = true;
            buffer.SetInputLabelInteriorColor();
            buffer.ClearBorderAllButTop();

            weights.GetTopRightCell().Offset[-2,-1].SetInvisibleRangeName(PolicyExcelMatrix.BasisRangeName);
            var basisRange = PolicyExcelMatrix.GetProfileBasisRange();
            basisRange.Locked = false;
            basisRange.SetInputDropdownInteriorColor();
            basisRange.SetInputFontColor();
            PolicyExcelMatrix.SetProfileBasesInWorksheet();
            PolicyExcelMatrix.SetProfileBasisInWorksheet();
            PolicyExcelMatrix.ImplementProfileBasis();
        }

        public override void ConvertToFlat()
        {
            if (!Validate()) return;
            
            var bodyRange = PolicyExcelMatrix.GetBodyRange();
            bodyRange.Clear();
            bodyRange.Offset[0, 1].Resize[1, bodyRange.Columns.Count - 3].EntireColumn.Delete();

            PolicyExcelMatrix.Dimension = new FlatPolicyProfileDimension(PolicyExcelMatrix);
            PolicyExcelMatrix.Reformat();
        }

        public override void ConvertToLimitBySir()
        {
            if (!Validate()) return;

            var bodyRange = PolicyExcelMatrix.GetBodyRange();
            bodyRange.Clear();

            PolicyExcelMatrix.Dimension = new LimitBySirPolicyProfileDimension(PolicyExcelMatrix);
            PolicyExcelMatrix.Reformat();
        }

        public override void ConvertToSirByLimit()
        {
            throw new System.NotImplementedException();
        }

        public override Range GetBodyHeaderRange()
        {
            return GetInputRange().GetTopLeftCell();
        }

        public override Range GetInputRange()
        {
            return PolicyExcelMatrix.RangeName.GetRange();
        }
    }
}