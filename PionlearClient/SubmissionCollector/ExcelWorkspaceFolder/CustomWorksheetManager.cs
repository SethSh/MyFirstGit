using System;
using System.Windows.Forms;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class CustomWorksheetManager
    {
        public static void AddWorksheet(IWorkbookLogger logger)
        {
            try
            {
                using (new WorkbookUnprotector())
                {
                    var lastSheet = Globals.ThisWorkbook.Sheets[Globals.ThisWorkbook.Sheets.Count];
                    Globals.ThisWorkbook.Worksheets.Add(After: lastSheet);
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Add worksheet failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static void MoveWorksheetRight(IWorkbookLogger logger)
        {
            try
            {
                if (Globals.ThisWorkbook.IsSelectedWorksheetOwnedByUser())
                    Globals.ThisWorkbook.MoveSelectedWorksheetRight();
                else
                    MessageHelper.Show("Can only move a worksheet owner by the user", MessageType.Stop);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Move worksheet failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static void MoveWorksheetLeft(IWorkbookLogger logger)
        {
            try
            {
                if (Globals.ThisWorkbook.IsSelectedWorksheetOwnedByUser())
                    Globals.ThisWorkbook.MoveSelectedWorksheetLeft();
                else
                    MessageHelper.Show("Can only move a worksheet owner by the user", MessageType.Stop);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Move worksheet failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static void RenameWorksheet(IWorkbookLogger logger)
        {
            try
            {
                if (Globals.ThisWorkbook.IsSelectedWorksheetOwnedByUser())
                    Globals.ThisWorkbook.RenameSelectedWorksheet();
                else
                    MessageHelper.Show("Can only rename a worksheet owner by the user", MessageType.Stop);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Rename worksheet failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static void DeleteWorksheet(IWorkbookLogger logger)
        {
            try
            {
                if (Globals.ThisWorkbook.IsSelectedWorksheetOwnedByUser())
                {
                    var confirmationResponse = MessageHelper.ShowWithYesNo("Are you sure you want to delete this worksheet?");
                    if (confirmationResponse == DialogResult.No) return;

                    Globals.ThisWorkbook.DeleteSelectedWorksheet();
                    return;
                }

                MessageHelper.Show("Can only delete a worksheet owner by the user", MessageType.Stop);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Delete worksheet failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}