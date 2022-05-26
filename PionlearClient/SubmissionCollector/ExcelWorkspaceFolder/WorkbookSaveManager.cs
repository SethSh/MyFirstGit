namespace SubmissionCollector.ExcelWorkspaceFolder
{
    public static class WorkbookSaveManager
    {
        public static bool IsSavedInitially
        {
            get
            {
                var workbook = Globals.ThisWorkbook;
                return !string.IsNullOrEmpty(workbook.Path);
            }
        }

        public static bool IsReadOnly
        {
            get
            {
                var workbook = Globals.ThisWorkbook;
                return workbook.ReadOnly;
            }
        }

        public static void Save()
        {
            var workbook = Globals.ThisWorkbook;
            workbook.Save();
        }
    }
}
