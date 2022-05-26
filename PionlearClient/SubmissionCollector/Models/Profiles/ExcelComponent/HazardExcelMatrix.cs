using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.View.Forms;
using IModel = PionlearClient.Model.IModel;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public class HazardExcelMatrix : MultipleOccurrenceProfileExcelMatrix, IRangeTransposable
    {
        public HazardExcelMatrix(int segmentId) : base(segmentId)
        {
            InterDisplayOrder = 10;
        }

        public HazardExcelMatrix()
        {
                
        }

        public override string FriendlyName => BexConstants.HazardProfileName;
        
        public override string ExcelRangeName => ExcelConstants.HazardProfileRangeName;

        public override string SublinesRangeName => string.Empty;
        public override string SublinesHeaderRangeName => string.Empty;

        public ISubline Subline => this.Single();

        public override void ModifyForChangeInSublines(int sublineCount)
        {
            var segment = GetSegment();
            var multipleOccurrenceSegmentExcelMatrices = segment.ExcelMatrices.OfType<MultipleOccurrenceSegmentExcelMatrix>().ToList();
            var maximumSublineCount = multipleOccurrenceSegmentExcelMatrices.Max(c => c.Count);
            MoveRangesWhenSublinesChange(maximumSublineCount);
        }

        public override int ComponentId
        {
            get => this.Single().Code;
            set { }
        }
        
        public override Range GetSublinesRange()
        {
            throw new NotImplementedException();
        }

        public override Range GetSublinesHeaderRange()
        {
            throw new NotImplementedException();
        }

        [JsonIgnore]
        public IList<HazardDistributionItemPlus> Items { set; get; }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maxDisplayOrder = GetSegment().HazardProfiles.Max(x => x.IntraDisplayOrder);
                
                if (IntraDisplayOrder != maxDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {FriendlyName.ToLower()} farther to the right.", MessageType.Stop);
                return false;
            }
        }
        public override bool IsOkToMoveLeft
        {
            get
            {
                var minDisplayOrder = GetSegment().HazardProfiles.Min(x => x.IntraDisplayOrder);

                if (IntraDisplayOrder != minDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {BexConstants.HazardProfileName.ToLower()} farther to the left.",
                    MessageType.Stop);
                return false;
            }
        }
        
        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = this.GetParent().Name;
            headerRange.SetToDefaultFont();

            var basisRange = GetProfileBasisRange();
            basisRange.Locked = false;
            basisRange.SetInputDropdownInteriorColor();
            basisRange.SetInputFontColor();
            SetProfileBasesInWorksheet();

            var bodyRange = GetBodyRange();
            bodyRange.ClearBorders();
            bodyRange.SetBorderToOrdinary();

            var bodyHeaderRange = GetBodyHeaderRange();
            bodyHeaderRange.ClearBorders();
            bodyHeaderRange.GetTopLeftCell().Locked = true;
            bodyHeaderRange.GetTopLeftCell().SetInputLabelColor();
            bodyHeaderRange.SetBorderToOrdinary();
            if (IsTransposed)
            {
                bodyHeaderRange.GetTopLeftCell().AlignRight();
            }
            else
            {
                bodyHeaderRange.GetTopLeftCell().AlignLeft();
                bodyHeaderRange.GetTopLeftCell().Value = "Name";
                bodyHeaderRange.GetTopRightCell().AlignRight();
            }

            var labelRange = GetInputLabelRange();
            labelRange.Locked = true;
            labelRange.SetInputLabelColor();
            if (IsTransposed)
            {
                labelRange.AlignRight();
            }
            else
            {
                labelRange.AlignLeft();
            }

            var inputRange = GetInputRange();
            inputRange.AlignRight();
            inputRange.SetInputColor();

            ImplementProfileBasis();

            bodyRange.GetColumn(0).ColumnWidth = 25;
            bodyRange.GetColumn(1).ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;
        }

        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(1, 1);
        }

        public override Range GetInputLabelRange()
        {
            return !IsTransposed ? RangeName.GetRangeSubset(1, 0).GetColumn(0) : RangeName.GetRangeSubset(0, 1).GetRow(0);
        }

        public override StringBuilder Validate()
        {
            var validations = IsTransposed ? ValidateTransposed() : ValidateOrdinary();

            var needToNormalize = ProfileFormatter.RequiresNormalization || 
                !ProfileFormatter.RequiresNormalization && Items.Sum(alloc => alloc.Value).IsEpsilonEqualToOne();
            if (needToNormalize) Items.Normalize();

            return validations;
        }
        
        private StringBuilder ValidateOrdinary()
        {
            var validation = new StringBuilder();
            Items = new List<HazardDistributionItemPlus>();

            var inputRange = GetInputRange();

            var names = GetInputLabelRange().GetContent();
            var weightsFromExcel = inputRange.GetContent();
            var weights = weightsFromExcel.ForceContentToDoubles();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();
            const int weightColumn = 0;
            var rowCount = names.GetLength(0);
            var columnLetter = (startColumn + weightColumn).GetColumnLetter();

            for (var row = 0; row < rowCount; row++)
            {
                var name = names[row, weightColumn].ToString();
                var weightFromExcel = weightsFromExcel[row, weightColumn];
                var weight = weights[row, weightColumn];
                
                if (weightFromExcel == null) continue;

                var rowNumber = row + startRow;
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, rowNumber);

                if (double.IsNaN(weight))
                {
                    var message = $"Enter {BexConstants.HazardProfileName.ToLower()} {weightLabel} in {addressLocation}" +
                                  $": <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                    continue;
                }

                var hazardId = HazardCodesFromBex.ReferenceData.Single(x => x.SubLineOfBusinessCode == Subline.Code && x.Name == name).Id;
                var item = new HazardDistributionItemPlus
                {
                    HazardId = hazardId, 
                    Value = weight,
                    Location = addressLocation
                };
                Items.Add(item);
            }

            return validation;
        }

        private StringBuilder ValidateTransposed()
        {
            var validation = new StringBuilder();
            Items = new List<HazardDistributionItemPlus>();

            var inputRange = GetInputRange();

            var names = GetInputLabelRange().GetFirstRow().GetContent();
            var weightsFromExcel = inputRange.GetContent();
            var weights = weightsFromExcel.ForceContentToDoubles();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();
            const int weightRow = 0;
            var columnCount = names.GetLength(1);
            var absoluteRow = startRow;

            for (var column = 0; column < columnCount; column++)
            {
                var name = names[weightRow, column].ToString();
                var weightFromExcel = weightsFromExcel[weightRow, column];
                var weight = weights[weightRow, column];
                
                if (weightFromExcel == null) continue;

                var columnNumber = startColumn + column;
                var columnLetter = columnNumber.GetColumnLetter();
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);

                if (double.IsNaN(weight))
                {
                    var message = $"Enter {BexConstants.HazardProfileName.ToLower()} {weightLabel} in {addressLocation}" +
                                  $": <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                    continue;
                }

                var item = new HazardDistributionItemPlus
                {
                    HazardId = HazardCodesFromBex.ReferenceData.Single(x => x.SubLineOfBusinessCode == Subline.Code && x.Name == name).Id,
                    Value = weight,
                    Location = addressLocation
                };
                Items.Add(item);
            }
            
            return validation;
        }
        
        public override Range GetBodyRange()
        {
            return !IsTransposed ? RangeName.GetRangeSubset(1, 0) : RangeName.GetRangeSubset(0, 1);
        }

        public override Range GetBodyHeaderRange()
        {
            return !IsTransposed ? RangeName.GetRange().GetRow(0) : RangeName.GetRange().GetColumn(0);
        }
        
        public static string GetRangeName(int segmentId, int sublineCode)
        {
            return GetRangeName(segmentId, sublineCode, ExcelConstants.HazardProfileRangeName);
        }

        public static string GetBasisRangeName(int segmentId, int sublineCode)
        {
            return GetBasisRangeName(segmentId, sublineCode, ExcelConstants.HazardProfileRangeName);
        }

        public static string GetHeaderRangeName(int segmentId, int sublineCode)
        {
            return GetHeaderRangeName(segmentId, sublineCode, ExcelConstants.HazardProfileRangeName);
        }

        public bool IsTransposed { get; set; }

        public virtual IModel GetParent()
        {
            return GetSegment().HazardProfiles.Single(x => x.ComponentId == ComponentId);
        }

        public override void SetColumnVisibility(ISubline subline)
        {
            if (subline.Code == 0)
            {
                HideColumns();
                return;
            }

            if (Subline.GetType() == subline.GetType())
            {
                ShowColumns();
            }
            else
            {
                HideColumns();
            }
        }

        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().HazardProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1).ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().HazardProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1).ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }
    }
}
