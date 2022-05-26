using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter
{
    [JsonObject]
    public class LimitBySirPolicyProfileDimension : BasePolicyProfileDimension
    {
        //              attachment 1   attachment 2
        //                  1k              5k
        //              -------------   -------------
        //limit 1   1m |    30%                 20%
        //limit 2   2m |    15%                 35%
        public LimitBySirPolicyProfileDimension(PolicyExcelMatrix policyExcelMatrix) : base(policyExcelMatrix)
        {

        }

        public override Range WeightRange => PolicyExcelMatrix.GetInputRange().GetRangeSubset(1, 1);

        public override void Reformat()
        {
            PolicyExcelMatrix.GetSublinesRange().BorderAround2();
            PolicyExcelMatrix.GetSublinesHeaderRange().BorderAround2();
            PolicyExcelMatrix.GetBodyRange().Offset[-1, 0].BorderAround2();

            var bodyRange = PolicyExcelMatrix.GetBodyRange();
            bodyRange.ClearBorders();

            const int bufferRowCount = 1;
            var limitCount = bodyRange.Rows.Count - 1 - bufferRowCount;
            var sirCount = bodyRange.Columns.Count - 1;

            var header = PolicyExcelMatrix.GetBodyHeaderRange();
            var limits = bodyRange.Offset[1, 0].Resize[limitCount, 1];
            var sirs = bodyRange.Offset[0, 1].Resize[1, sirCount];
            var weights = bodyRange.Offset[1, 1].Resize[limitCount, sirCount];
            var buffer = bodyRange.GetLastRow();

            header.SetInputLabelInteriorColor();
            header.Value2 = $"Lim {RangeExtensions.DownArrow} SIR {RangeExtensions.RightArrow}";
            header.Locked = true;

            limits.NumberFormat = FormatExtensions.WholeNumberFormat;
            limits.SetBorderAroundToOrdinary();
            limits.SetBorderBottomToResizable();
            
            sirs.NumberFormat = FormatExtensions.WholeNumberFormat;
            sirs.SetBorderAroundToOrdinary();
            sirs.SetBorderRightToResizable();
            
            weights.NumberFormat = FormatExtensions.PercentFormat;
            weights.SetBorderRightToResizable();
            weights.SetBorderBottomToResizable();
            
            buffer.Locked = true;
            buffer.SetInputLabelInteriorColor();
            buffer.ClearBorderAllButTop();

            weights.GetTopRightCell().Offset[-2, -1].SetInvisibleRangeName(PolicyExcelMatrix.BasisRangeName);
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
            bodyRange.Offset[0,1].Resize[1, bodyRange.Columns.Count - 3].EntireColumn.Delete();
            
            PolicyExcelMatrix.Dimension = new FlatPolicyProfileDimension(PolicyExcelMatrix);
            PolicyExcelMatrix.Reformat();
        }

        public override void ConvertToLimitBySir()
        {
            throw new System.NotImplementedException();
        }

        public override void ConvertToSirByLimit()
        {
            if (!Validate()) return;

            var bodyRange = PolicyExcelMatrix.GetBodyRange();
            bodyRange.Clear();

            PolicyExcelMatrix.Dimension = new SirByLimitPolicyProfileDimension(PolicyExcelMatrix);
            PolicyExcelMatrix.Reformat();
        }

        public override Range GetInputRange()
        {
            return PolicyExcelMatrix.RangeName.GetRange();
        }

        public override Range GetBodyHeaderRange()
        {
            return GetInputRange().GetTopLeftCell();
        }
    }
}