using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.Models.DataComponents
{
    public interface IProfileExcelMatrix : ISegmentExcelMatrix , IRangeSumable
    {
        IProfileFormatter ProfileFormatter { get; set; }
        
        string BasisRangeName { get; }
        Range GetProfileBasisRange();
        bool ValidateBasis();
    }
    
}