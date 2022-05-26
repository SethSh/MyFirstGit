using System.Text;
using Microsoft.Office.Interop.Excel;

namespace SubmissionCollector.Models.DataComponents
{
    public interface IExcelMatrix 
    {
        string FriendlyName { get; }
        string FullName { get; }
        string ExcelRangeName { get; }
        string RangeName { get; set; }
        int ColumnStart { get; }
        void Reformat();
        void ToggleEstimate();
        Range GetInputRange();
        Range GetInputLabelRange();
        StringBuilder Validate();
        bool HasData { get; }
    }
}