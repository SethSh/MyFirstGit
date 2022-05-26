using System.Linq;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models.Segment.DataComponents
{
    internal class SegmentExcelComponentIdentifier
    {
        public ISegment Segment { get; set; }
        public ISegmentExcelMatrix ExcelMatrix { get; set; }
        public Range TopLeftSelectedCell { get; set; }
        public Range SelectedRange { get; set; }
        private bool IsQuiet { get; set; }

        public bool Validate(bool isQuiet = false)
        {
            IsQuiet = isQuiet;
            var worksheetValidator = new SegmentWorksheetValidator();
            if (!worksheetValidator.Validate(IsQuiet)) return false;

            Segment = worksheetValidator.Segment;
            SelectedRange = worksheetValidator.SelectedRange;
            TopLeftSelectedCell = SelectedRange.GetTopLeftCell();
            return ValidateDataComponent();
        }
        

        private bool ValidateDataComponent()
        {
            var rangeNames = Segment.ExcelMatrices.Where(x => !(x is UmbrellaExcelMatrix)).Select(x => x.RangeName).ToList();
            if (Segment.IsUmbrella) rangeNames.Add(Segment.UmbrellaExcelMatrix.RangeName);

            var rangeName = string.Empty;
            foreach (var item in rangeNames)
            {
                if (!item.ContainsRange(TopLeftSelectedCell)) continue;
                rangeName = item;
                break;
            }

            ExcelMatrix = Segment.ExcelMatrices.SingleOrDefault(x => x.RangeName == rangeName);
            if (ExcelMatrix != null) return true;

            if (!IsQuiet) MessageHelper.Show(@"The selection must be within an input range", MessageType.Stop);
            return false;
        }
    }
}
