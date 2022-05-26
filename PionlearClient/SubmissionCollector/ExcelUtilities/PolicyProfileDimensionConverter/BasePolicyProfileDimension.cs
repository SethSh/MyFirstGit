using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using SubmissionCollector.Enums;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter
{
    [JsonObject]
    public abstract class BasePolicyProfileDimension : IPolicyProfileDimension
    {
        protected BasePolicyProfileDimension(PolicyExcelMatrix policyExcelMatrix)
        {
            PolicyExcelMatrix = policyExcelMatrix;
        }


        public PolicyExcelMatrix PolicyExcelMatrix { get; }

        [JsonIgnore]
        public abstract Range WeightRange { get; }

        public virtual bool CheckIsAllNull()
        {
            var rangeContent = PolicyExcelMatrix.GetInputRange().GetContent();
            var rowCount = rangeContent.GetLength(0);
            var columnCount = rangeContent.GetLength(1);

            var rowContent = rangeContent.GetRow(0);
            for (var column = 1; column < columnCount; column++)
            {
                if (rowContent[column] != null) return false;
            }

            for (var row = 1; row < rowCount; row++)
            {
                rowContent = rangeContent.GetRow(row);
                if (!rowContent.AllNull()) return false;
            }
            return true;
        }

        public abstract void Reformat();

        public bool Validate()
        {
            if (CheckIsAllNull()) return true;

            MessageHelper.Show($"Can only change {PionlearClient.BexConstants.PolicyProfileName.ToLower()} dimension when range is empty", MessageType.Stop);
            return false;
        }

        public abstract void ConvertToFlat();
        public abstract void ConvertToLimitBySir();
        public abstract void ConvertToSirByLimit();

        public abstract Range GetInputRange();
        public abstract Range GetBodyHeaderRange();
    }
}