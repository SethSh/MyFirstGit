using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal abstract class BaseMultiOccurenceExcelMatrixHelper<T> : BaseExcelMatrixHelper<T> where T : IExcelComponent
    {
        protected BaseMultiOccurenceExcelMatrixHelper(ISegment segment) : base(segment)
        {
        
        }
        
        public void CreateRanges()
        {
            var adjacentRange = GetAdjacentRange();

            var counter = 0;
            foreach (var excelMatrix in ExcelComponents.Select(ec => ec.CommonExcelMatrix).Where(ec => ec is MultipleOccurrenceSegmentExcelMatrix))
            {
                var columnOffset = IsPreviousRangeAdjacent ? ColumnCount : ColumnCountPlusOne;
                var anchorRange = adjacentRange.Offset[0, 1].Offset[0, columnOffset * counter++];
                InsertRanges(anchorRange, (MultipleOccurrenceSegmentExcelMatrix)excelMatrix);
            }

            DeleteTemplateRange(Segment, ExcelRangeName, !IsPreviousRangeAdjacent);
        }

        public void ModifyRanges()
        {
            #region add new ranges to worksheet
            var adjacentRange = GetAdjacentRange();

            var componentsToAdd = ExcelComponents.Where(ec => !ec.CommonExcelMatrix.RangeName.ExistsInWorkbook() && ec.CommonExcelMatrix is MultipleOccurrenceSegmentExcelMatrix);
            var counter = 0;
            foreach (var excelMatrix in componentsToAdd.Select(prof => prof.CommonExcelMatrix))
            {
                var columnOffset = IsPreviousRangeAdjacent ? ColumnCount : ColumnCountPlusOne; 
                var anchorRange = adjacentRange.Offset[0, 1].Offset[0, columnOffset * counter++];
                InsertRanges(anchorRange, (MultipleOccurrenceSegmentExcelMatrix)excelMatrix);
            }
            #endregion

            DeleteOrphanRanges(ExcelComponents.ToList(), $"segment{Segment.Id}.{ExcelRangeName}", !IsPreviousRangeAdjacent);
            
            #region ensure sort order is correct
            if (ExcelComponents.Count() <= 1) return;

            var maxDisplayOrder = ExcelComponents.Max(x => x.IntraDisplayOrder);
            for (var i = 0; i < maxDisplayOrder; i++)
            {
                var orderInExcel = ExcelComponents
                    .OrderBy(component => component.CommonExcelMatrix.RangeName.GetTopRightCell().Column)
                    .Select(profile => profile.IntraDisplayOrder).ToList();
                var index = orderInExcel.IndexOf(i);

                var steps = index - i;
                if (steps <= 0) continue;

                var excelMatrix = ExcelComponents.Select(component => component.CommonExcelMatrix).Where(mat => mat is MultipleOccurrenceSegmentExcelMatrix).SingleOrDefault(matrix => matrix.IntraDisplayOrder == i);
                ((MultipleOccurrenceSegmentExcelMatrix) excelMatrix)?.MoveLeft(steps);
            }
            #endregion
        }


        private static void DeleteTemplateRange(ISegment segment, string rangeName, bool hasBufferColumn = true)
        {
            var templateRangeName = $"{ExcelConstants.SegmentRangeName}{segment.Id}.{rangeName}";
            if (templateRangeName.ExistsInWorkbook())
            {
                var templateRange = templateRangeName.GetRange();
                Globals.ThisWorkbook.Names.Item(templateRangeName).Delete();
                if (hasBufferColumn)
                {
                    templateRange.AppendColumn().EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                }
                else
                {
                    templateRange.EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                }

            }

            var templateSublineRangeName = $"{templateRangeName}.{BexConstants.SublineName.ToLower()}s";
            if (templateSublineRangeName.ExistsInWorkbook())
            {
                Globals.ThisWorkbook.Names.Item(templateSublineRangeName).Delete();
            }

        }

        public abstract void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix);

        protected int GetWorksheetSublineRowCount()
        {
            var a = Segment.ProspectiveExposureAmountExcelMatrix.GetInputRange().Row;
            var b = Segment.SublineExcelMatrix.HeaderRangeName.GetTopLeftCell().Row;
            var templateSublineRowCount = b - a + 1 - ExcelConstants.InBetweenRowCount;
            return templateSublineRowCount;
        }

        private static  void DeleteOrphanRanges(IList<T> excelComponents, string prefix, bool appendColumn = true)
        {
            var anyDeleted = false;

            //segment0.exposureSet0 for example
            //escape d means number
            //+ means 0, 10, 100 counts as a number
            //$ means comes at end of string
            var pattern = prefix + @"\d+$";
            var regex = new Regex(pattern);
            var allRangeNames = RangeExtensions.GetMatchingRangeNames(regex);
            var modelRangeNames = excelComponents.Select(excelComponent => excelComponent.CommonExcelMatrix.RangeName);

            foreach (var rangeName in allRangeNames.Except(modelRangeNames))
            {
                anyDeleted = true;
                if (appendColumn)
                {
                    rangeName.GetRange().AppendColumn().EntireColumn.Delete();
                }
                else
                {
                    rangeName.GetRange().EntireColumn.Delete();
                    foreach (var excelComponent in excelComponents)
                    {
                        excelComponent.CommonExcelMatrix.Reformat();
                    }
                }
                rangeName.DeleteRangeName();
            }

            if (anyDeleted) WorkbookExtensions.RemoveErrorRanges();
        }
    }
}