using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models.Package
{
    internal class PackageWorksheetValidator
    {
        public Range SelectedRange { get; set; }
        public bool IsQuiet { get; set; }


        public bool Validate()
        {
            SelectedRange = Globals.ThisWorkbook.Application.Selection as Range;
            if (SelectedRange == null)
            {
                if (!IsQuiet) MessageHelper.Show(@"Can't recognized selected item as a range", MessageType.Stop);
                return false;
            }

            var worksheet = Globals.ThisWorkbook.GetSelectedWorksheet();
            if (worksheet == null)
            {
                if (!IsQuiet) MessageHelper.Show(@"Can't verify that a worksheet is selected", MessageType.Stop);
                return false;
            }

            return Globals.ThisWorkbook.ThisExcelWorkspace.Package.Worksheet.Name == worksheet.Name;
        }
    }
}
