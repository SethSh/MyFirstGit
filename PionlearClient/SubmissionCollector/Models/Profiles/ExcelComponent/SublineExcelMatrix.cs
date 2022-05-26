using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using MunichRe.Bex.ApiClient.CollectorApi;
using Newtonsoft.Json;
using PionlearClient;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public class SublineExcelMatrix : SingleOccurrenceProfileExcelMatrix
    {
        public SublineExcelMatrix(int segmentId) : base(segmentId)
        {
            
        }


        [JsonIgnore]
        public List<Allocation> Allocations { get; set; }

        public override string FriendlyName => BexConstants.SublineAllocationName;
        public override string ExcelRangeName => ExcelConstants.SublineAllocationRangeName;

        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetToDefaultFont();

            var bodyHeaderRange = GetBodyHeaderRange();
            bodyHeaderRange.Locked = true;
            bodyHeaderRange.SetInputLabelColor();
            bodyHeaderRange.SetBorderToOrdinary();
            

            var basisRange = GetProfileBasisRange();
            basisRange.Locked = false;
            basisRange.SetInputDropdownInteriorColor();
            basisRange.SetInputFontColor();
            SetProfileBasesInWorksheet();

            var labelRange = GetInputLabelRange();
            labelRange.Locked = true;
            labelRange.SetInputLabelColor();
            
            var bodyRange = GetBodyRange();
            bodyRange.SetBorderToOrdinary();
            
            ImplementProfileBasis();
        }

        public override Range GetInputRange()
        {
            return GetBodyRange().GetColumn(1);
        }

        public override Range GetInputLabelRange()
        {
            return GetBodyRange().GetColumn(0);
        }

        public override StringBuilder Validate()
        {
            var validations = new StringBuilder();

            ValidateBasis(validations);

            var segment = GetSegment();
            var inputRange = GetInputRange();
            var values = inputRange.GetContent();
            var valuesAsDoubles = values.ForceContentToDoubles();

            Allocations = new List<Allocation>();
            var row = 0;
            foreach (var item in segment.ToList())
            {
                var rowBaseOne = row + 1;
                var value = values[row, 0];
                var valueAsDouble = valuesAsDoubles[row, 0];
                if (value == null)
                {
                    validations.AppendLine($"Subline weight in row {rowBaseOne} can't be blank");
                }
                else if (double.IsNaN(valuesAsDoubles[row, 0]))
                {
                    validations.AppendLine($"Subline weight <{value}> in row {rowBaseOne} is not a number");
                }

                Allocations.Add(new Allocation
                {
                    Id = item.Code,
                    Value = valueAsDouble
                });
                row++;
            }
            
            var needToNormalize = ProfileFormatter.RequiresNormalization ||
                                  !ProfileFormatter.RequiresNormalization && Allocations.Sum(alloc => alloc.Value).IsEpsilonEqualToOne();
            if (needToNormalize) Allocations.Normalize();
            
            return validations;
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }

        public override Range GetBodyRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public override void InsertHeaderRows(Range prospectiveRange, int desiredTopRow)
        {
            const int inBetweenRowCount = ExcelConstants.InBetweenRowCount;
            var headerRange = GetHeaderRange();
            var headerRangeTopRow = headerRange.GetTopLeftCell().Row;

            var delta = desiredTopRow - headerRangeTopRow;
            var insertRange = headerRange.Offset[-inBetweenRowCount, 0].Resize[delta, headerRange.Columns.Count];

            var needFormatCorrection = insertRange.GetTopRightCell().Offset[-1, 0].Address == prospectiveRange.Address;
            insertRange.InsertRangeDown();

            //when the start of the insert is one row under the prospect exposure ... copies the darn formats
            if (needFormatCorrection)
            {
                var formatRange = insertRange.Offset[-delta, 0];
                formatRange.Locked = false;
                formatRange.ClearInteriorColor();
                formatRange.ClearFontColor();
            }
        }
    }
}
