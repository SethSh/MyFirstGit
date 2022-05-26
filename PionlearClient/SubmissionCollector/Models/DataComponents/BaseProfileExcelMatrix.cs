using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.Models.DataComponents
{
    public abstract class BaseProfileExcelMatrix : BaseSegmentExcelMatrix, IProfileExcelMatrix
    {
        protected BaseProfileExcelMatrix() 
        {

        }

        protected BaseProfileExcelMatrix(int segmentId) : base(segmentId)
        {
                
        }

        public IProfileFormatter ProfileFormatter { get; set; }

        public string BasisRangeName => $"{ExcelConstants.SegmentRangeName}{SegmentId}.{ExcelRangeName}.{ExcelConstants.ProfileBasisRangeName}";

        public void SetProfileBasesInWorksheet()
        {
            var profileBases = ProfileBasisFromBex.NamesInOrder.ToArray();
            var profileBasisInDropdown = string.Join(", ", profileBases);

            var range = GetBodyHeaderRange().GetTopRightCell();
            range.AlignRight();
            range.Validation.Delete();

            var segment = GetSegment();
            var worksheet = segment.WorksheetManager.Worksheet;
            if (worksheet.ProtectContents)
            {
                worksheet.UnprotectInterface();
                range.Validation.Add(XlDVType.xlValidateList, Formula1: profileBasisInDropdown);
                worksheet.ProtectInterface();
            }
            else
            {
                range.Validation.Add(XlDVType.xlValidateList, Formula1: profileBasisInDropdown); 
            }
        }

        public void SetProfileBasisInWorksheet()
        {
            GetProfileBasisRange().Value2 = ProfileFormatter is PercentProfileFormatter
                ? ProfileBasisFromBex.ReferenceData.Single(p => p.Id.ToString() == 1.ToString()).Name
                : ProfileBasisFromBex.ReferenceData.Single(p => p.Id.ToString() == 2.ToString()).Name;
        }
        
        public virtual void ImplementProfileBasis()
        {
            var inputRange = GetInputRange();
            var sumRange = GetSumRange();
            var sumFormula = $"=Sum({GetInputRange().Address})";

            ProfileFormatter.FormatDataRange(inputRange);
            ProfileFormatter.WriteSumFormulaToRange(sumRange, sumFormula);
        }

        public Range GetProfileBasisRange()
        {
            return BasisRangeName.GetTopRightCell();
        }

        public Range GetSumRange()
        {
            return HeaderRangeName.GetTopRightCell();
        }
        
        public bool ValidateBasis(StringBuilder sb)
        {
            var profileBasisRange = GetProfileBasisRange();
            var acceptableValues = ProfileBasisFromBex.NamesInOrder.ToList();

            if (profileBasisRange.Value2 != null && acceptableValues.Contains(profileBasisRange.Resize[1, 1].Value2.ToString()))
            {
                return true;
            }

            sb.AppendLine($"{FriendlyName} basis not recognized");
            return false;
        }

        public bool ValidateBasis()
        {
            var ignore = new StringBuilder();
            return ValidateBasis(ignore);
        }


    }
}