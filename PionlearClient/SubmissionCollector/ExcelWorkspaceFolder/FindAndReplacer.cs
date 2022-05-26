using System;
using Excel = Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class FindAndReplacer
    {
        internal static void FindAndReplace(IWorkbookLogger logger)
        {
            try
            {
                FindAndReplace();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Find and Replace failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void FindAndReplace()
        {
            var selectWorksheet = Globals.ThisWorkbook.GetSelectedWorksheet();
            if (selectWorksheet == null)
            {
                MessageHelper.Show("Can't recognize selected worksheet", MessageType.Stop);
                return;
            }

            var selectedRange = Globals.ThisWorkbook.GetSelectedRange();
            if (selectedRange == null)
            {
                MessageHelper.Show("Can't recognize selected range", MessageType.Stop);
                return;
            }

            var isProtected = selectWorksheet.ProtectContents;

            if (isProtected)
            {
                if (selectedRange.Rows.Count == 1 && selectedRange.Columns.Count == 1)
                {
                    MessageHelper.Show("Select a non [1 x 1] range", MessageType.Stop);
                    return;
                }

                foreach (Excel.Range range in selectedRange)
                {
                    if (!(bool) range.Locked) continue;

                    MessageHelper.Show("Can't find and replace unless the entire selected range is unlocked", MessageType.Stop);
                    return;
                }
            }

            var app = (Microsoft.Office.Interop.Excel.Application) Globals.ThisWorkbook.Application;
            app.Dialogs[Excel.XlBuiltInDialog.xlDialogFormulaReplace].Show();
        }
    }
}