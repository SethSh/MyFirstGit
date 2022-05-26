using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter
{
    [JsonObject]
    public interface IPolicyProfileDimension
    {
        PolicyExcelMatrix PolicyExcelMatrix { get; }
        Range WeightRange { get; }
        bool CheckIsAllNull();
        void Reformat();
        bool Validate();
        void ConvertToFlat();
        void ConvertToLimitBySir();
        void ConvertToSirByLimit();
        Range GetInputRange();
        Range GetBodyHeaderRange();
    }
}