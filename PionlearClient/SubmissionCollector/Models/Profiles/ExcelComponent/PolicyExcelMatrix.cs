using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.CollectorClientPlus;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.View.Forms;
using IModel = PionlearClient.Model.IModel;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public sealed class PolicyExcelMatrix : MultipleOccurrenceProfileExcelMatrix, IRangeResizable
    {
        public PolicyExcelMatrix(int segmentId, int componentId) : base(segmentId)
        {
            Dimension = new FlatPolicyProfileDimension(this);
            ComponentId = componentId;
            InterDisplayOrder = 20;
        }

        public PolicyExcelMatrix()
        {

        }

        public override string FullName
        {
            get
            {
                var segment = GetSegment();
                var profile = (PolicyProfile)GetParent();
                if (!segment.IsUmbrella || !profile.UmbrellaType.HasValue)
                {
                    return base.FullName;
                }

                var fullName = base.FullName;
                return fullName.Replace(FriendlyName, profile.Name);
            }
        }

        public override string FriendlyName => BexConstants.PolicyProfileName;
        public override string ExcelRangeName => ExcelConstants.PolicyProfileRangeName;

        public IPolicyProfileDimension Dimension { get; set; }

        [JsonIgnore]
        public IList<PolicyDistributionItemPlus> Items { get; set; }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maxDisplayOrder = GetSegment().PolicyProfiles.Max(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != maxDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {FriendlyName.ToLower()} farther to the right.", MessageType.Stop);
                return false;
            }
        }
        public override bool IsOkToMoveLeft
        {
            get
            {
                var minDisplayOrder = GetSegment().PolicyProfiles.Min(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != minDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {FriendlyName.ToLower()} farther to the left.",
                    MessageType.Stop);
                return false;
            }
        }
        
        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetToDefaultFont();
            
            //inserting in first row issue of having validation in inserted cell
            var inputRange = GetInputRange();
            inputRange.Validation.Delete();

            var bodyRange = GetBodyRange();
            bodyRange.ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;

            Dimension.Reformat();
        }

        public void ReformatBorderTop()
        {
            if (Dimension is FlatPolicyProfileDimension)
            {
                RangeName.GetRange().GetRow(1).SetBorderTopToOrdinary();
            }
        }

        public void ReformatBorderBottom()
        {
            if (Dimension is FlatPolicyProfileDimension)
            {
                RangeName.GetRange().GetLastRow().SetBorderBottomToResizable();
            }
        }

        public override Range GetInputRange()
        {
            return Dimension.GetInputRange();
        }

        public override Range GetInputLabelRange()
        {
            throw new NotImplementedException();
        }

        public override StringBuilder Validate()
        {
            var validations = new StringBuilder();

            if (Dimension is FlatPolicyProfileDimension) validations = ValidateFlat();
            if (Dimension is LimitBySirPolicyProfileDimension) validations = ValidateLimitBySir();
            if (Dimension is SirByLimitPolicyProfileDimension) validations = ValidateSirByLimit();

            var needToNormalize = ProfileFormatter.RequiresNormalization ||
                                  !ProfileFormatter.RequiresNormalization && Items.Sum(alloc => alloc.Value).IsEpsilonEqualToOne();
            if (needToNormalize) Items.Normalize();

            return validations;
        }

        public IModel GetParent()
        {
            return GetSegment().PolicyProfiles.Single(x => x.ComponentId == ComponentId);
        }

        public override Range GetBodyHeaderRange()
        {
            return Dimension.GetBodyHeaderRange();
        }

        public override Range GetBodyRange()
        {
            return GetInputRange();
        }
       
        public static void HideAllColumns(int segmentId)
        {
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                GetRangeName(segmentId, i).GetRange().AppendColumn().HideColumns();
            }
        }
        
        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.PolicyProfileRangeName}{componentId}";
        }

        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().PolicyProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1).ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().PolicyProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1).ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }
        
        private StringBuilder ValidateFlat()
        {
            var validation = new StringBuilder();
            Items = new List<PolicyDistributionItemPlus>();

            var inputRange = GetInputRange();
            var content = inputRange.GetContent();
            var input = content.ForceContentToDoubles();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();

            var rowCount = input.GetLength(0);
            var columnCount = input.GetLength(1);

            const int limitColumn = 0;
            const int sirColumn = 1;
            const int weightColumn = 2;

            var columnLetters = new Dictionary<int, string>();
            for (var column = 0; column < columnCount; column++)
            {
                columnLetters.Add(column, (startColumn + column).GetColumnLetter());
            }

            for (var row = 0; row < rowCount; row++)
            {
                var limitFromExcel = content[row, 0];
                var sirFromExcel = content[row, 1];
                var weightFromExcel = content[row, 2];

                if (limitFromExcel == null && sirFromExcel == null && weightFromExcel == null) continue;

                var limit = input[row, limitColumn];
                var sir = sirFromExcel != null ? input[row, sirColumn] : new double?();
                var weight = input[row, weightColumn];

                var rowNumber = row + startRow;
                var columnLetter = columnLetters[limitColumn];
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                if (limitFromExcel == null)
                {
                    validation.AppendLine($"Enter {BexConstants.LimitName.ToLower()} in {addressLocation}");
                }
                else if (double.IsNaN(limit))
                {
                    validation.AppendLine($"Enter {BexConstants.LimitName.ToLower()} in {addressLocation}: <{limitFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }
                
                columnLetter = columnLetters[sirColumn];
                addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                if (sirFromExcel != null && double.IsNaN(sir.Value))
                {
                    validation.AppendLine($"Enter {BexConstants.SirAttachmentName.ToLower()} in {addressLocation}: <{sirFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }

                columnLetter = columnLetters[weightColumn];
                addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                if (weightFromExcel == null)
                {
                    validation.AppendLine($"Enter {weightLabel} in {addressLocation}");
                }
                else if (double.IsNaN(weight))
                {
                    validation.AppendLine($"Enter {weightLabel} in {addressLocation}: <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }

                Items.Add(new PolicyDistributionItemPlus
                {
                    Limit = limit, 
                    Attachment = sir, 
                    Value = weight,
                    Location = $"row {rowNumber}"
                });
            }

            return validation;
        }

        private StringBuilder ValidateLimitBySir()
        {
            var validation = new List<string>();
            Items = new List<PolicyDistributionItemPlus>();

            var inputRange = GetInputRange();
            var content = inputRange.GetContent();
            var input = content.ForceContentToDoubles();
            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var rowCount = content.GetLength(0);
            var columnCount = content.GetLength(1);

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();

            var columnLetters = new Dictionary<int, string>();
            for (var column = 0; column < columnCount; column++)
            {
                columnLetters.Add(column, (startColumn + column).GetColumnLetter());
            }
            const int limitColumn = 0;
            const int sirRow = 0;

            for (var row = 1; row < rowCount; row++)
            {
                var limitFromExcel = content[row, 0];
                if (limitFromExcel == null) continue;

                var limit = input[row, 0];
                if (!double.IsNaN(limit)) continue;

                var absoluteRow = row + startRow;
                var columnLetter = columnLetters[limitColumn];
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);
                validation.Add($"Enter {BexConstants.LimitName.ToLower()} in {addressLocation}: <{limitFromExcel}> {BexConstants.NotRecognizedAsANumber}");
            }

            for (var column = 1; column < columnCount; column++)
            {
                var sirFromExcel = content[0, column];
                if (sirFromExcel == null) continue;

                var sir = input[0, column];
                if (!double.IsNaN(sir)) continue;

                var absoluteRow = startRow;
                var columnLetter = columnLetters[column];
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);
                validation.Add($"Enter {BexConstants.SirAttachmentName.ToLower()}  in {addressLocation}:" +
                               $" <{sirFromExcel}> {BexConstants.NotRecognizedAsANumber}");
            }

            for (var row = 1; row < rowCount; row++)
            {
                var absoluteRow = row + startRow;
                
                for (var column = 1; column < columnCount; column++)
                {
                    var weightFromExcel = content[row, column];
                    if (weightFromExcel == null) continue;

                    var weight = input[row, column];
                    if (!double.IsNaN(weight)) continue;

                    var columnLetter = columnLetters[column];
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);
                    validation.Add($"{addressLocation} {weightLabel} <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }
            }

            for (var row = 1; row < rowCount; row++)
            {
                var rowNumber = row + startRow;

                for (var column = 1; column < columnCount; column++)
                {
                    var limit = input[row, limitColumn];
                    var sir = input[sirRow, column];
                    var weight = input[row, column];

                    var limitFromExcel = content[row, limitColumn];
                    var sirFromExcel = content[sirRow, column];
                    var weightFromExcel = content[row, column];
                    
                    if (weightFromExcel == null) continue;
                    var failsValidation = false;

                    if (limitFromExcel == null)
                    {
                        var columnLetter = columnLetters[limitColumn];
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                        validation.Add($"Enter {BexConstants.LimitName.ToLower()} in {addressLocation}");
                        failsValidation = true;
                    }

                    if (sirFromExcel == null)
                    {
                        var columnLetter = columnLetters[column];
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                        validation.Add($"Enter {BexConstants.SirAttachmentName.ToLower()} in {addressLocation}");
                        failsValidation = true;
                    }
                    if (failsValidation) continue;

                    Items.Add(new PolicyDistributionItemPlus   
                    {
                        Limit = limit, 
                        Attachment = sir, 
                        Value = weight,
                        Location = $"row {rowNumber}"
                    });
                }
            }

            var distinctValidation = validation.Distinct();
            var sb = new StringBuilder();
            distinctValidation.ForEach(x => sb.AppendLine(x));

            return sb;
        }

        private StringBuilder ValidateSirByLimit()
        {
            var validation = new List<string>();
            Items = new List<PolicyDistributionItemPlus>();

            var inputRange = GetInputRange();
            var content = inputRange.GetContent();
            var input = content.ForceContentToDoubles();
            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var rowCount = content.GetLength(0);
            var columnCount = content.GetLength(1);

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();


            var columnLetters = new Dictionary<int, string>();
            for (var column = 0; column < columnCount; column++)
            {
                columnLetters.Add(column, (startColumn + column).GetColumnLetter());
            }
            const int sirColumn = 0;
            const int limitRow = 0;

            for (var row = 1; row < rowCount; row++)
            {
                var sirFromExcel = content[row, sirColumn];
                if (sirFromExcel == null) continue;

                var sir = input[row, sirColumn];
                if (!double.IsNaN(sir)) continue;

                var rowNumber = row + startRow;
                var absoluteColumn = columnLetters[sirColumn];
                var address = $"{BexConstants.RangeName.ToLower()} {absoluteColumn}{rowNumber}";
                validation.Add($"Enter {BexConstants.SirAttachmentName.ToLower()} in {address}: <{sirFromExcel}> {BexConstants.NotRecognizedAsANumber}");
            }

            for (var column = 1; column < columnCount; column++)
            {
                var limitFromExcel = content[limitRow, column];
                if (limitFromExcel == null) continue;

                var limit = input[limitRow, column];
                if (double.IsNaN(limit))
                {
                    var absoluteRow = startRow;
                    var columnLetter = columnLetters[column];
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);
                    validation.Add($"Enter {BexConstants.LimitName.ToLower()} in {addressLocation}: <{limitFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }
            }

            for (var row = 1; row < rowCount; row++)
            {
                var absoluteRow = row + startRow;
                
                for (var column = 1; column < columnCount; column++)
                {
                    var weightFromExcel = content[row, column];
                    if (weightFromExcel == null) continue;

                    var weight = input[row, column];
                    if (!double.IsNaN(weight)) continue;

                    var columnLetter = columnLetters[column];
                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);
                    validation.Add($"{addressLocation} {weightLabel} <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                }
            }

            for (var row = 1; row < rowCount; row++)
            {
                var rowNumber = row + startRow;

                for (var column = 1; column < columnCount; column++)
                {
                    var sir = input[row,sirColumn];
                    var limit = input[limitRow, column];
                    var weight = input[row, column];

                    var sirFromExcel = content[row, sirColumn];
                    var limitFromExcel = content[limitRow, column];
                    var weightFromExcel = content[row, column];

                    if (weightFromExcel == null) continue;
                    var failsValidation = false;
                    
                    if (sirFromExcel == null)
                    {
                        var columnLetter = columnLetters[sirColumn];
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                        validation.Add($"Enter {BexConstants.SirAttachmentName.ToLower()} in {addressLocation}");
                        failsValidation = true;
                    }

                    if (limitFromExcel == null)
                    {
                        var columnLetter = columnLetters[column];
                        var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);
                        validation.Add($"Enter {BexConstants.LimitName.ToLower()} in {addressLocation}");
                        failsValidation = true;
                    }

                    if (failsValidation) continue;

                    Items.Add(new PolicyDistributionItemPlus
                    {
                        Limit = limit,
                        Attachment = sir,
                        Value = weight,
                        Location = $"row {rowNumber}"
                    });
                }
            }

            var distinctValidation = validation.Distinct();
            var sb = new StringBuilder();
            distinctValidation.ForEach(x => sb.AppendLine(x));

            return sb;
        }

        public override void ImplementProfileBasis()
        {
            var inputRange = Dimension.WeightRange;
            var sumRange = GetSumRange();
            var sumFormula = $"=Sum({Dimension.WeightRange.Address})";

            ProfileFormatter.FormatDataRange(inputRange);
            ProfileFormatter.WriteSumFormulaToRange(sumRange, sumFormula);
        }

        public override void ToggleEstimate()
        {
            IsEstimate = !IsEstimate;

            var inputRange = GetInputRange().RemoveLastRow();
            if (Dimension is FlatPolicyProfileDimension)
            {
                inputRange.RemoveLastRow();
            }
            SetInputInteriorColorContemplatingEstimate(inputRange);
        }

        public static string GetBasisRangeName(int segmentId, int componentId)
        {
            return $"segment{segmentId}.{ExcelConstants.PolicyProfileRangeName}{componentId}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return GetHeaderRangeName(segmentId, componentId, ExcelConstants.PolicyProfileRangeName);
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return GetSublinesRangeName(segmentId, componentId, ExcelConstants.PolicyProfileRangeName);
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return GetSublinesHeaderRangeName(segmentId, componentId, ExcelConstants.PolicyProfileRangeName);
        }
    }
}
