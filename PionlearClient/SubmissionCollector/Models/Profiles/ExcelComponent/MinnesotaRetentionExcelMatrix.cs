using System;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public sealed class MinnesotaRetentionExcelMatrix : SingleOccurrenceProfileExcelMatrix, IRangeResizable
    {
        public MinnesotaRetentionExcelMatrix(int segmentId) : base(segmentId)
        {
            InterDisplayOrder = 36;
        }

        public override string FriendlyName => BexConstants.MinnesotaRetentionName;
        public override string ExcelRangeName => ExcelConstants.MinnesotaRetentionRangeName;


        [JsonIgnore]
        public long? RetentionId { get; set; }

        [JsonIgnore]
        public double RetentionValue { get; set; }

        
        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = BexConstants.MinnesotaRetentionName;
            headerRange.SetToDefaultFont();

            var labelRange = GetInputLabelRange();
            labelRange.SetInputLabelFormat();
            labelRange.AlignLeft();
            labelRange.Value2 = "Retention";

            var inputRange = GetInputRange();
            inputRange.NumberFormat = FormatExtensions.WholeNumberFormat;
            inputRange.SetInputDropdownInteriorColor();
            
            labelRange.Union(inputRange).SetBorderAroundToOrdinary();

            labelRange.ColumnWidth = ExcelConstants.StandardColumnWidth;
            inputRange.ColumnWidth = ExcelConstants.StandardColumnWidth;
            inputRange.Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;
        }

        
        public override Range GetInputRange()
        {
            return RangeName.GetTopRightCell();
        }

        public override Range GetInputLabelRange()
        {
            return RangeName.GetTopLeftCell();
        }

        public override Range GetBodyHeaderRange()
        {
            throw new ArgumentException();
        }
        
        public override StringBuilder Validate()
        {
            var validation = new StringBuilder();

            var retentionValue = GetInputRange().Value;
            if (retentionValue == null)
            {
                var message = "Minnesota Retention can't be blank";
                validation.AppendLine(message);
                return validation;
            }

            if (!double.TryParse(retentionValue.ToString(), out double d))
            {
                var message = "Minnesota Retention not recognized as a number";
                validation.AppendLine(message);
            }

            var id = MinnesotaRetentionsFromBex.ReferenceData
                .SingleOrDefault(item => item.RetentionAmount == Convert.ToInt64(retentionValue))?.Id;

            RetentionId = id;
            RetentionValue = d;

            return validation;
        }

        public static string GetRangeName(int segmentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.MinnesotaRetentionRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId)
        {
            return $"{GetRangeName(segmentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetBasisRangeName(int segmentId)
        {
            return $"{GetRangeName(segmentId)}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public override Range GetBodyRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public void ReformatBorderTop()
        {
            GetInputRange().GetFirstRow().SetBorderTopToOrdinary();
        }

        public override void ToggleEstimate()
        {
            IsEstimate = !IsEstimate;

            var inputRange = GetInputRange();
            SetInputDropdownInteriorColorContemplatingEstimate(inputRange);
        }

        

        public void ReformatBorderBottom()
        {
            RangeName.GetRange().GetLastRow().SetBorderBottomToOrdinary();
        }
        
    }
}
