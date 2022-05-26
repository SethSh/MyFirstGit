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
    public sealed class ProtectionClassExcelMatrix : MultipleOccurrenceProfileExcelMatrix, IRangeTransposable
    {
        public ProtectionClassExcelMatrix(int segmentId) : base(segmentId)
        {
            InterDisplayOrder = 34;
        }

        public ProtectionClassExcelMatrix()
        {

        }

        [JsonIgnore]
        public IList<ProtectionClassDistributionItemPlus> Items { get; set; }

        public override string FriendlyName => BexConstants.ProtectionClassProfileName;
        public override string ExcelRangeName => ExcelConstants.ProtectionClassProfileRangeName;

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

        public override Range GetInputRange()
        {
            var body = GetBodyRange();
            return !IsTransposed ? body.GetLastColumn() : body.GetLastRow();
        }

        public override Range GetInputLabelRange()
        {
            var body = GetBodyRange();
            return !IsTransposed ? body.GetFirstColumn() : body.GetFirstRow();
        }

        public override StringBuilder Validate()
        {
            var validations = IsTransposed ? ValidateTransposed() : ValidateOrdinary();

            var needToNormalize = ProfileFormatter.RequiresNormalization ||
                                  !ProfileFormatter.RequiresNormalization && Items.Sum(alloc => alloc.Weight).IsEpsilonEqualToOne();
            if (needToNormalize) Items.Normalize();

            return validations;
        }

        public override Range GetBodyHeaderRange()
        {
            return !IsTransposed ? RangeName.GetRange().GetRow(0) : RangeName.GetRange().GetColumn(0);
        }

        public override Range GetBodyRange()
        {
            return !IsTransposed ? RangeName.GetRangeSubset(1, 0) : RangeName.GetRangeSubset(0, 1);
        }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maxDisplayOrder = GetSegment().ProtectionClassProfiles.Max(x => x.IntraDisplayOrder);
                if (IntraDisplayOrder != maxDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {FriendlyName.ToLower()} farther to the right.", MessageType.Stop);
                return false;
            }
        }
        public override bool IsOkToMoveLeft
        {
            get
            {
                var minDisplayOrder = GetSegment().ProtectionClassProfiles.Min(x => x.IntraDisplayOrder);
                if (IntraDisplayOrder != minDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {FriendlyName.ToLower()} farther to the left.", MessageType.Stop);
                return false;
            }
        }
        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().ProtectionClassProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1).ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().ProtectionClassProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1).ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }

        public IModel GetParent()
        {
            return GetSegment().ProtectionClassProfiles.Single(x => x.ComponentId == ComponentId);
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

        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = this.GetParent().Name;

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

        public bool IsTransposed { get; set; }

        private StringBuilder ValidateOrdinary()
        {
            var validation = new StringBuilder();
            Items = new List<ProtectionClassDistributionItemPlus>();

            var inputRange = GetInputRange();

            var names = GetInputLabelRange().GetFirstColumn().GetContent();
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
                var name = names[row, weightColumn];
                var weightFromExcel = weightsFromExcel[row, weightColumn];
                var weight = weights[row, weightColumn];

                if (weightFromExcel == null) continue;

                var absoluteRow = row + startRow;
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);

                if (double.IsNaN(weight))
                {
                    var message = $"Enter {BexConstants.ProtectionClassProfileName.ToLower()} {weightLabel} in {addressLocation}" +
                                  $": <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                    continue;
                }

                var code = ProtectionClassCodesFromBex.ReferenceData.Single(pc => pc.SubLineOfBusinessCode == Subline.Code && pc.Name == name.ToString()).Id;
                var item = new ProtectionClassDistributionItemPlus
                {
                    ProtectionClassId = Convert.ToInt32(code),
                    Weight = weight,
                    Location = addressLocation
                };
                Items.Add(item);
            }

            return validation;
        }

        private StringBuilder ValidateTransposed()
        {
            var validation = new StringBuilder();
            Items = new List<ProtectionClassDistributionItemPlus>();

            var inputRange = GetInputRange();

            var names = GetInputLabelRange().GetFirstRow().GetContent();
            var weightsFromExcel = inputRange.GetContent();
            var weights = weightsFromExcel.ForceContentToDoubles();

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();
            const int weightRow = 0;
            var columnCount = names.GetLength(1);

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var absoluteRow = startRow;
            for (var column = 0; column < columnCount; column++)
            {
                var name = names[weightRow, column];
                var weightFromExcel = weightsFromExcel[weightRow, column];
                var weight = weights[weightRow, column];

                if (weightFromExcel == null) continue;

                var absoluteColumn = startColumn + column;
                var columnLetter = absoluteColumn.GetColumnLetter();
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);

                if (double.IsNaN(weight))
                {
                    var message = $"Enter {BexConstants.ProtectionClassProfileName.ToLower()} {weightLabel} in {addressLocation}" +
                                  $": <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                    continue;
                }

                var code = ProtectionClassCodesFromBex.ReferenceData.Single(pc => pc.SubLineOfBusinessCode == Subline.Code && pc.Name == name.ToString()).Id;
                var item = new ProtectionClassDistributionItemPlus
                {
                    ProtectionClassId = Convert.ToInt32(code),
                    Weight = weight,
                    Location = addressLocation
                };
                Items.Add(item);
            }

            return validation;
        }

        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.ProtectionClassProfileRangeName}{componentId}";
        }

        public static string GetBasisRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.ProtectionClassProfileRangeName}{componentId}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return GetSublinesRangeName(segmentId, componentId, ExcelConstants.ProtectionClassProfileRangeName);
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return GetSublinesHeaderRangeName(segmentId, componentId, ExcelConstants.ProtectionClassProfileRangeName);
        }
    }


}
