using System;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Segment.DataComponents
{
    public class ProspectiveExposureAmountExcelMatrix : SingleOccurrenceSegmentExcelMatrix
    {
        public ProspectiveExposureAmountExcelMatrix(int segmentId) : base(segmentId)
        {
            
        }

        public override string FriendlyName => BexConstants.ProspectiveExposureAmountName;
        public override string ExcelRangeName => ExcelConstants.ProspectiveExposureAmountRangeName;

        [JsonIgnore]
        public double Item { get; set; }

        public override void Reformat()
        {
            base.Reformat();

            var range = RangeName.GetRange();
            range.ClearBorders();
            range.SetBorderToOrdinary();

            var labelRange = GetInputLabelRange();
            labelRange.Locked = true;
            labelRange.SetInputLabelColor();
            
            var inputRange = GetInputRange();
            inputRange.NumberFormat = FormatExtensions.WholeNumberFormat;
        }

        public override Range GetInputRange()
        {
            return RangeName.GetRange().GetTopRightCell();
        }

        public override Range GetInputLabelRange()
        {
            return RangeName.GetRange().GetTopLeftCell();
        }
        
        public override Range GetBodyRange()
        {
            throw new NotImplementedException();
        }

        public override Range GetBodyHeaderRange()
        {
            throw new NotImplementedException();
        }

        public override StringBuilder Validate()
        {
            var value = GetInputRange().GetContent();
            var validation = new StringBuilder();
            if (value[0, 0] == null)
            {
                validation.AppendLine($"Enter a number in {BexConstants.ProspectiveExposureAmountName.ToLower()}");
                Item = double.NaN;
            }
            else
            {
                var valueAsDouble = value.ForceContentToDoubles();
                Item = valueAsDouble[0, 0];
                if (double.IsNaN(Item))
                {
                    validation.AppendLine($"Enter a number in {BexConstants.ProspectiveExposureAmountName.ToLower()}: '{value[0, 0]} isn't a number");
                }
            }

            return validation;
        }

        public override void MoveRangesWhenSublinesChange(int sublineCount)
        {
            //don't do the base behavior
        }
    }
}
