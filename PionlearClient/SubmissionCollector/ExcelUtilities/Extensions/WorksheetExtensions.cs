using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ExcelUtilities.Extensions
{
    public static class WorksheetExtensions
    {
        private static readonly Application ExcelApplication = Globals.ThisWorkbook.Application;

        public static void SelectFirstCell(this Worksheet worksheet)
        {
            ExcelApplication.Goto(worksheet.Range["A1"]);
        }

        public static void UnprotectInterface(this Worksheet worksheet)
        {
            worksheet.Unprotect(Password: BexConstants.WorkbookPassword);
        }

        public static void ProtectInterface(this Worksheet worksheet)
        {
            worksheet.Protect(Password: BexConstants.WorkbookPassword,
                UserInterfaceOnly: true,
                DrawingObjects: false,
                Contents: true,
                Scenarios: false,
                AllowFormattingCells: true,
                AllowFormattingColumns: true,
                AllowFormattingRows: true,
                AllowInsertingHyperlinks: true,
                AllowSorting: true,
                AllowFiltering: true,
                AllowUsingPivotTables: true);
        }

        internal static Worksheet GetWorksheet(this string worksheetName)
        {
            return Globals.ThisWorkbook.Worksheets.Cast<Worksheet>()
                .FirstOrDefault(item => item.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase));
        }

        public static Worksheet GetWorksheetWithNullTrap(this string worksheetName)
        {
            foreach (Worksheet item in Globals.ThisWorkbook.Worksheets)
            {
                if (item.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return item;
                }
            }

            throw new ArgumentOutOfRangeException($"No worksheet exists named {worksheetName}");
        }

        public static IPackage GetPackage(this Worksheet worksheet)
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            return package.Worksheet.Name == worksheet.Name ? package : null;
        }

        public static ISegment GetSegment(this Worksheet worksheet)
        {
            return Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.SingleOrDefault(x => x.WorksheetManager.Worksheet.Name == worksheet.Name);
        }

        public static void LockNecessaryCells(this Worksheet worksheet)
        {
            worksheet.UsedRange.Locked = false;
            worksheet.UsedRange.LockNecessaryCells();
        }

        public static void SetVisibleToHidden(this Worksheet worksheet)
        {
            using (new ExcelEventDisabler())
            {
                using (new WorkbookUnprotector())
                {
                    worksheet.Visible = XlSheetVisibility.xlSheetHidden;
                }
            }
        }
    }
}
