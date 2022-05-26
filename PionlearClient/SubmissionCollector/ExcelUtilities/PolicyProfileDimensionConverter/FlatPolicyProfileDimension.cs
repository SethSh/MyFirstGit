using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter
{
    public class FlatPolicyProfileDimension : BasePolicyProfileDimension
    {
        //limit     attachment      percent
        //1m        1k              .25
        //2m        1k              .15
        //1m        5k              .05
        //2m        5k              .05

        public FlatPolicyProfileDimension(PolicyExcelMatrix policyExcelMatrix) : base(policyExcelMatrix)
        {
                
        }
        
        public override bool CheckIsAllNull()
        {
            var rangeContent = PolicyExcelMatrix.GetInputRange().GetContent();
            var rowCount = rangeContent.GetLength(0);
            for (var row = 0; row < rowCount; row++)
            {
                var rowContent = rangeContent.GetRow(row);
                if (!rowContent.AllNull()) return false;
            }
            return true;
        }

        public override void Reformat()
        {
            var sublinesHeaderRange = PolicyExcelMatrix.GetSublinesHeaderRange();
            sublinesHeaderRange.SetSublineHeaderFormat();
            sublinesHeaderRange.ClearContents();
            sublinesHeaderRange.GetTopLeftCell().Value = $"{BexConstants.PolicyProfileName} {BexConstants.SublineName}s";

            var headerRange = PolicyExcelMatrix.GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = BexConstants.PolicyProfileName;
            headerRange.SetToDefaultFont();
            headerRange.GetRangeSubset(0, 1).RemoveLastColumn().ClearContents();

            var bodyHeaderRange = PolicyExcelMatrix.GetBodyHeaderRange();
            bodyHeaderRange.ClearBorders();
            bodyHeaderRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            bodyHeaderRange.SetBorderToOrdinary();

            bodyHeaderRange.Resize[1, 2].Value2 = new[] { BexConstants.LimitName, BexConstants.SirAttachmentName};
            bodyHeaderRange.Resize[1, 2].Locked = true;
            bodyHeaderRange.Resize[1, 2].SetInputLabelColor();
            
            bodyHeaderRange.GetTopRightCell().SetInvisibleRangeName(PolicyExcelMatrix.BasisRangeName);
            var basisRange = PolicyExcelMatrix.GetProfileBasisRange();
            basisRange.Locked = false;
            basisRange.SetInputDropdownInteriorColor();
            basisRange.SetInputFontColor();
            PolicyExcelMatrix.SetProfileBasesInWorksheet();
            PolicyExcelMatrix.SetProfileBasisInWorksheet();

            var inputRange = PolicyExcelMatrix.GetInputRange();
            var inputRangeWithoutBuffer = inputRange.RemoveLastRow();
            inputRangeWithoutBuffer.SetBorderAroundToResizable();
            inputRangeWithoutBuffer.Resize[inputRangeWithoutBuffer.Rows.Count, 2].NumberFormat = FormatExtensions.WholeNumberFormat;
            
            PolicyExcelMatrix.ImplementProfileBasis();
            
            var inputRangeBuffer = inputRange.GetLastRow();
            inputRangeBuffer.ClearContents();
            inputRangeBuffer.Locked = true;
            inputRangeBuffer.SetInputLabelInteriorColor();
            inputRangeBuffer.ClearBorderAllButTop();
        }

        public override Range WeightRange => PolicyExcelMatrix.GetInputRange().GetLastColumn().RemoveLastRow();

        public override void ConvertToFlat()
        {
            throw new System.NotImplementedException();
        }

        public override void ConvertToLimitBySir()
        {
            if (!Validate()) return;
            
            const int desiredSirCount = 10;
            const int desiredColumnCount = desiredSirCount + 1;  //one additional column for limits

            var bodyRange = PolicyExcelMatrix.GetBodyRange();
            var headerRange = PolicyExcelMatrix.GetBodyHeaderRange();
            bodyRange.Union(headerRange).Clear();
            bodyRange.Offset[0,1].Resize[1, desiredColumnCount - bodyRange.Columns.Count].InsertColumnsToRight();
            
            PolicyExcelMatrix.Dimension = new LimitBySirPolicyProfileDimension(PolicyExcelMatrix);
            PolicyExcelMatrix.Reformat();
        }

        public override void ConvertToSirByLimit()
        {
            if (!Validate()) return;
            
            const int desiredLimitCount = 10;
            const int desiredColumnCount = desiredLimitCount + 1; //one additional column for sirs

            var bodyRange = PolicyExcelMatrix.GetBodyRange();
            var headerRange = PolicyExcelMatrix.GetBodyHeaderRange();
            bodyRange.Union(headerRange).Clear();
            bodyRange.Offset[0, 1].Resize[1, desiredColumnCount - bodyRange.Columns.Count].InsertColumnsToRight();
            
            PolicyExcelMatrix.Dimension = new  SirByLimitPolicyProfileDimension(PolicyExcelMatrix);
            PolicyExcelMatrix.Reformat();
        }

        public override Range GetInputRange()
        {
            return PolicyExcelMatrix.RangeName.GetRangeSubset(1, 0);
        }

        public override Range GetBodyHeaderRange()
        {
            return PolicyExcelMatrix.RangeName.GetRange().GetRow(0);
        }
    }
}