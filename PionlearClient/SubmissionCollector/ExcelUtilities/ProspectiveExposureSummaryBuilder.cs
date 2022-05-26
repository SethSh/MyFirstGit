using System;
using System.Linq;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;

namespace SubmissionCollector.ExcelUtilities
{
    internal class ProspectiveExposureSummaryBuilder
    {
        private const string SummaryRangeName = "submission.prospectiveExposureSummary";
        private const string PremiumName = "Premium";
        internal void Build()
        {
            using (new ExcelScreenUpdateDisabler())
            {
                var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
                var range = SummaryRangeName.GetRangeSubset(1, 0);
                var basisRange = range.GetTopRightCell();
                var itemsRange = range.GetRangeSubset(1, 0).RemoveLastRow();
                var totalRange = range.GetBottomRightCell();

                if (!package.Segments.Any())
                {
                    itemsRange.ClearContents();
                    totalRange.ClearContents();
                    basisRange.ClearContents();
                    return;
                }

                var segments = package.Segments.OrderBy(s => s.DisplayOrder).ToList();

                #region insert/delete rows
                if (segments.Count > itemsRange.Rows.Count)
                {
                    var delta = segments.Count - itemsRange.Rows.Count;
                    itemsRange.GetTopLeftCell().Offset[1,0].Resize[delta, itemsRange.Columns.Count].InsertRangeDown();
                }
                else if (itemsRange.Rows.Count >1 && itemsRange.Rows.Count > segments.Count)
                {
                    var delta = itemsRange.Rows.Count - segments.Count;
                    itemsRange.GetTopLeftCell().Offset[1, 0].Resize[delta, itemsRange.Columns.Count].DeleteRangeUp();
                }
                #endregion

                #region write into body
                range = SummaryRangeName.GetRangeSubset(1, 0);
                itemsRange = range.GetRangeSubset(1, 0).RemoveLastRow();
                itemsRange.SetBorderAroundToOrdinary();
                var segmentNameAddresses = segments.Select(x => $"={x.HeaderRangeName.GetRange().GetBottomRightCell().Address[External:true]}").ToArray();
                var exposureAddresses = segments.Select(x => $"={x.ProspectiveExposureAmountExcelCell}").ToArray();
                var obj = new object[segments.Count, 2];
                for (var row = 0; row < segments.Count; row++)
                {
                    obj[row, 0] = segmentNameAddresses[row];
                    obj[row, 1] = exposureAddresses[row];
                }
                itemsRange.Formula = obj;
                #endregion

                #region exposure base 
                var exposureBasisIds = package.Segments.Select(x => Convert.ToInt16(x.ProspectiveExposureBasis)).Distinct();
                var exposureBasisNames = exposureBasisIds.Select(basisId => ExposureBasisFromBex.GetExposureBasisName(basisId)).ToList();

                if (exposureBasisNames.Count > 1 && exposureBasisNames.All(x => x.Contains(PremiumName)))
                {
                    basisRange.GetTopLeftCell().Value2 = PremiumName;
                }
                else
                {
                    basisRange.GetTopLeftCell().Value2 = exposureBasisNames.Count == 1 ? exposureBasisNames.Single() : "Multiple";
                }
                #endregion

                #region total 
                var formula = $"= Sum({itemsRange.GetLastColumn().Address})";
                totalRange.Formula = formula;
                #endregion
            }
        }
    }
}
