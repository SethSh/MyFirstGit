using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace SubmissionCollector.ExcelUtilities.Extensions
{
    public static class WorkbookExtensions
    {
        public static void RemoveErrorRanges()
        {
            foreach (var item in Globals.ThisWorkbook.Names.Cast<Name>().Where(item => item.RefersTo.ToString().Contains("#REF")))
            {
                item.Delete();
            }
        }
    }
}
