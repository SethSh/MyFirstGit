using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Models.DataComponents
{
    //I couldn't double inherit, so I duplicated BaseProfileExcelComponent
    public abstract class MultipleOccurrenceProfileExcelMatrix : MultipleOccurrenceSegmentExcelMatrix, IProfileExcelMatrix
    {
        protected MultipleOccurrenceProfileExcelMatrix(int segmentId) : base (segmentId)
        {
            var userPreferences = UserPreferences.ReadFromFile();
            ProfileFormatter = ProfileFormatterFactory.Create(userPreferences.ProfileBasisId);
        }

        protected MultipleOccurrenceProfileExcelMatrix()
        {
                
        }

        public override string FullName
        {
            get
            {
                var list = new List<ISubline>();
                foreach (var subline in this)
                {
                    list.Add(subline);
                }

                var names = GetLineSublineConcatenatedNames(list);
                return $"{FriendlyName} {string.Join(", ", names)}";
            }
        }

        public virtual void ImplementProfileBasis()
        {
            var inputRange = GetInputRange();
            var sumRange = GetSumRange();
            var sumFormula = $"=Sum({GetInputRange().Address})";
            sumRange.AlignRight();

            ProfileFormatter.FormatDataRange(inputRange);
            ProfileFormatter.WriteSumFormulaToRange(sumRange, sumFormula);
        }

        public virtual Range GetSumRange()
        {
            return HeaderRangeName.GetTopRightCell();
        }

        public IProfileFormatter ProfileFormatter { get; set; }
        
        public string BasisRangeName => $"{ExcelConstants.SegmentRangeName}{SegmentId}.{ExcelRangeName}{ComponentId}.{ExcelConstants.ProfileBasisRangeName}";

        public void SetProfileBasesInWorksheet()
        {
            var profileBases = ProfileBasisFromBex.NamesInOrder.ToArray();
            var profileBasisInDropdown = string.Join(", ", profileBases);
            var range = GetProfileBasisRange();
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
            IProfileFormatterHandler profileFormatterHandler;
            if (ProfileFormatter is PercentProfileFormatter)
            {
                profileFormatterHandler = new PercentProfileFormatterHandler();
            }
            else if (ProfileFormatter is PremiumProfileFormatter)
            {
                profileFormatterHandler = new PremiumProfileFormatterHandler();
            }
            else
            {
                const string message = "Can't find profile formatter";
                throw new ArgumentOutOfRangeException(message);
            }
            
            GetProfileBasisRange().Value2 = ProfileBasisFromBex.ReferenceData.Single(p => p.Id == profileFormatterHandler.ProfileBasisId).Name;
        }
        
        public Range GetProfileBasisRange()
        {
            return BasisRangeName.GetTopRightCell();
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

        public override void MoveRight()
        {
            var profiles = GetProfileSiblings();
            var excelMatrix = ((BaseProfile)profiles.Single(profile => ((BaseProfile)profile).IntraDisplayOrder == IntraDisplayOrder + 1)).CommonExcelMatrix;
            var columnShift = excelMatrix.RangeName.GetRange().AppendColumn().Columns.Count;

            var range = RangeName.GetRange().AppendColumn();
            var columnCount = range.Columns.Count;

            range.Offset[0, columnCount + columnShift].InsertColumnsToRight();
            range.EntireColumn.Cut(range.Offset[0, columnCount + columnShift].EntireColumn);
            range.Offset[0, -(columnCount + columnShift)].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        public override void MoveLeft(int steps = 1)
        {
            //when called from ribbon display order is ok to use
            //when called from subline wizard, display order is not ok to use because model and worksheet are out of sync
            var excelMatrices = GetProfileSiblingExcelMatrices();

            var columnStarts = excelMatrices.Select(p => p.ColumnStart).ToList();
            columnStarts.Sort();

            var myColumnStart = ColumnStart;
            var otherColumnStart = columnStarts[columnStarts.IndexOf(myColumnStart) - steps];
            var columnShift = myColumnStart - otherColumnStart;

            var range = RangeName.GetRange().AppendColumn();
            var columnCount = range.Columns.Count;

            range.Offset[0, -columnShift].InsertColumnsToRight();
            range.EntireColumn.Cut(range.Offset[0, -(columnShift + columnCount)].EntireColumn);
            range.Offset[0, columnShift + columnCount].EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        private IEnumerable<IProfile> GetProfileSiblings()
        {
            var segment = GetSegment();
            return segment.Profiles.Where(p => GetType().Name == ((BaseProfile)p).CommonExcelMatrix.GetType().Name);
        }

        private IEnumerable<ISegmentExcelMatrix> GetProfileSiblingExcelMatrices()
        {
            return GetProfileSiblings().Select(p => ((BaseProfile)p).CommonExcelMatrix);
        }
    }
}
