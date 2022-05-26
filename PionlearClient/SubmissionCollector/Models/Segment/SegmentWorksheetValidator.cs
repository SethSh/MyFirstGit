using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models.Segment
{
    internal class SegmentWorksheetValidator
    {
        public Range ActiveCell { get; set; }
        public Range SelectedRange { get; set; }
        public ISegment Segment { get; set; }
    
        public bool Validate(bool isQuiet = false)
        {
            ActiveCell = Globals.ThisWorkbook.Application.ActiveCell;
            SelectedRange = Globals.ThisWorkbook.Application.Selection as Range;
            if (SelectedRange == null)
            {
                if (!isQuiet) MessageHelper.Show(@"Selected item not recognized as a range", MessageType.Warning);
                return false;
            }

            var worksheet = Globals.ThisWorkbook.GetSelectedWorksheet();
            if (worksheet == null)
            {
                if (!isQuiet) MessageHelper.Show(@"Can't verify that a worksheet is selected", MessageType.Warning);
                return false;
            }

            Segment = worksheet.GetSegment();
            if (Segment != null) return true;

            if (!isQuiet) MessageHelper.Show(@"The active worksheet isn't a submission segment worksheet", MessageType.Warning);
            return false;
        }
        
    }
}
