using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using MunichRe.Bex.ApiClient.CollectorApi;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public class UmbrellaExcelMatrix : SingleOccurrenceProfileExcelMatrix
    {
        public UmbrellaExcelMatrix(int segmentId) : base(segmentId)
        {

        }

        [JsonIgnore] public List<Allocation> Allocations { get; set; }

        public override string FriendlyName => BexConstants.UmbrellaAllocationName;
        public override string ExcelRangeName => ExcelConstants.UmbrellaAllocationRangeName;

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

            var inputRange = GetInputRange();
            inputRange.SetInputColor();
            inputRange.Locked = false;
            
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
            Allocations = new List<Allocation>();

            if (GetSegment().ContainsAnyCommercialSublines)
            {
                ValidateBasis(validations);

                var names = GetInputLabelRange().GetContent();
                var weightsFromExcel = GetInputRange().GetContent();
                var weightsAsDoubles = weightsFromExcel.ForceContentToDoubles();

                for (var row = 0; row < names.GetLength(0); row++)
                {
                    var rowBaseOne = row + 1;
                    var name = names[row, 0].ToString();
                    var weightFromExcel = weightsFromExcel[row, 0];
                    var weightAsDouble = weightsAsDoubles[row, 0];

                    if (weightFromExcel == null)
                    {
                        validations.AppendLine($"{BexConstants.UmbrellaTypeName.ToStartOfSentence()}" +
                                               $" weight in row {rowBaseOne} can't be blank");
                    }
                    else if (double.IsNaN(weightAsDouble))
                    {
                        validations.AppendLine($"{BexConstants.UmbrellaTypeName.ToStartOfSentence()}" +
                                               $" weight '{weightAsDouble}' in row {rowBaseOne} is not a number");
                    }

                    var umbrellaCode = UmbrellaTypesFromBex.GetCode(name);
                    Allocations.Add(new Allocation
                    {
                        Id = umbrellaCode,
                        Value = weightAsDouble
                    });
                }

                var needToNormalize = ProfileFormatter.RequiresNormalization ||
                                      !ProfileFormatter.RequiresNormalization &&
                                      Allocations.Sum(alloc => alloc.Value).IsEpsilonEqualToOne();
                if (needToNormalize) Allocations.Normalize();
            }

            if (!GetSegment().ContainsAnyPersonalSublines) return validations;

            var personalCode = UmbrellaTypesFromBex.GetPersonalCode();
            Allocations.Add(new Allocation { Id = personalCode, Value = 1d });
            return validations;
        }


        public override Range GetBodyRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }

        public override void MoveRangesWhenSublinesChange(int sublineCount)
        {
            if (!RangeName.ExistsInWorkbook()) return;
            base.MoveRangesWhenSublinesChange(sublineCount);
        }
        
    }
}
