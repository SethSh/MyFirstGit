using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.View.Forms;
using IModel = PionlearClient.Model.IModel;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public sealed class StateExcelMatrix : MultipleOccurrenceProfileExcelMatrix, IRangeTransposable
    {
        public StateExcelMatrix(int segmentId, int componentId) : base(segmentId)
        {
            ComponentId = componentId;
            InterDisplayOrder = 30;
        }

        public StateExcelMatrix()
        {
        }

        public override string FriendlyName => BexConstants.StateProfileName;
        public override string ExcelRangeName => ExcelConstants.StateProfileRangeName;

        [JsonIgnore] public IList<StateDistributionItemPlus> Items { get; set; }

        public override bool IsOkToMoveRight
        {
            get
            {
                var maxDisplayOrder = GetSegment().StateProfiles.Max(x => x.IntraDisplayOrder);
                if (IntraDisplayOrder != maxDisplayOrder) return true;

                MessageHelper.Show($"Can't move the right-most {FriendlyName.ToLower()} farther to the right.", MessageType.Stop);
                return false;
            }
        }

        public override bool IsOkToMoveLeft
        {
            get
            {
                var minDisplayOrder = GetSegment().StateProfiles.Min(x => x.IntraDisplayOrder);
                if (IntraDisplayOrder != minDisplayOrder) return true;

                MessageHelper.Show($"Can't move the left-most {FriendlyName.ToLower()} farther to the left.", MessageType.Stop);
                return false;
            }
        }

        public bool IsTransposed { get; set; }

        public override void Reformat()
        {
            base.Reformat();

            var sublinesHeaderRange = GetSublinesHeaderRange();
            sublinesHeaderRange.SetSublineHeaderFormat();
            sublinesHeaderRange.ClearContents();
            sublinesHeaderRange.GetTopLeftCell().Value = $"{BexConstants.StateProfileName} {BexConstants.SublineName}s";

            var sublinesRange = GetSublinesRange();
            sublinesRange.SetSublineFormat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = BexConstants.StateProfileName;

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
                bodyHeaderRange.AlignRight();
            }
            else
            {
                bodyHeaderRange.AlignLeft();
                bodyHeaderRange.GetTopLeftCell().Value = "Abbrev";
                bodyHeaderRange.GetTopLeftCell().Offset[0, 1].Value = "Name";
                bodyHeaderRange.GetTopLeftCell().Offset[0, 1].Locked = true;
                bodyHeaderRange.GetTopLeftCell().Offset[0, 1].SetInputLabelColor();
                bodyHeaderRange.GetTopRightCell().AlignRight();
            }

            var labelRange = GetInputLabelRange();
            labelRange.Locked = true;
            labelRange.SetInputLabelColor();
            if (IsTransposed)
                labelRange.AlignRight();
            else
                labelRange.AlignLeft();

            var inputRange = GetInputRange();
            inputRange.AlignRight();
            inputRange.SetInputColor();

            ImplementProfileBasis();

            bodyRange.ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;
        }

        public override Range GetInputRange()
        {
            var body = GetBodyRange();
            return !IsTransposed ? body.GetColumn(2) : body.GetRow(2);
        }

        public override Range GetInputLabelRange()
        {
            var body = GetBodyRange();
            return !IsTransposed ? body.Resize[body.Rows.Count, 2] : body.Resize[2, body.Columns.Count];
        }

        public override StringBuilder Validate()
        {
            var validations = IsTransposed ? ValidateTransposed() : ValidateOrdinary();

            var needToNormalize = ProfileFormatter.RequiresNormalization ||
                                  !ProfileFormatter.RequiresNormalization && Items.Sum(alloc => alloc.Value).IsEpsilonEqualToOne();
            if (needToNormalize) Items.Normalize();

            return validations;
        }

        public override void SwitchDisplayOrderAfterMoveRight()
        {
            var otherExcelMatrix = GetSegment().StateProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder + 1).ExcelMatrix;

            IntraDisplayOrder++;
            otherExcelMatrix.IntraDisplayOrder--;
        }

        public override void SwitchDisplayOrderAfterMoveLeft()
        {
            var otherExcelMatrix = GetSegment().StateProfiles.Single(x => x.IntraDisplayOrder == IntraDisplayOrder - 1).ExcelMatrix;

            IntraDisplayOrder--;
            otherExcelMatrix.IntraDisplayOrder++;
        }

        public IModel GetParent()
        {
            return GetSegment().StateProfiles.Single(x => x.ComponentId == ComponentId);
        }


        public override Range GetBodyHeaderRange()
        {
            return !IsTransposed ? RangeName.GetRange().GetRow(0) : RangeName.GetRange().GetColumn(0);
        }

        public override Range GetBodyRange()
        {
            return !IsTransposed ? RangeName.GetRangeSubset(1, 0) : RangeName.GetRangeSubset(0, 1);
        }

        private StringBuilder ValidateOrdinary()
        {
            var validation = new StringBuilder();
            Items = new List<StateDistributionItemPlus>();

            var inputRange = GetInputRange();

            var abbreviations = GetInputLabelRange().GetFirstColumn().GetContent();
            var weightsFromExcel = inputRange.GetContent();
            var weights = weightsFromExcel.ForceContentToDoubles();

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();
            const int weightColumn = 0;
            var rowCount = abbreviations.GetLength(0);
            var columnLetter = (startColumn + weightColumn).GetColumnLetter();

            for (var row = 0; row < rowCount; row++)
            {
                var abbreviation = abbreviations[row, weightColumn];
                var weightFromExcel = weightsFromExcel[row, weightColumn];
                var weight = weights[row, weightColumn];

                if (weightFromExcel == null) continue;

                var absoluteRow = row + startRow;
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);

                if (double.IsNaN(weight))
                {
                    var message = $"Enter {BexConstants.StateProfileName.ToLower()} {weightLabel} in {addressLocation}" +
                                  $": <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                    continue;
                }

                //Todo figure out how to send CW now that ID isn't nullable
                var abbrevCode = abbreviation.ToString() == "CW"
                    ? new int()
                    : StateCodesFromBex.ReferenceData.Single(x => x.Abbreviation == abbreviation.ToString()).Id;

                var item = new StateDistributionItemPlus
                {
                    StateCode = abbrevCode,
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
            Items = new List<StateDistributionItemPlus>();

            var inputRange = GetInputRange();

            var abbreviations = GetInputLabelRange().GetFirstRow().GetContent();
            var weightsFromExcel = inputRange.GetContent();
            var weights = weightsFromExcel.ForceContentToDoubles();

            var weightLabel = GetProfileBasisRange().Value2.ToString().ToLower();
            const int weightRow = 0;
            var columnCount = abbreviations.GetLength(1);

            var topLeftCell = inputRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;

            var absoluteRow = startRow;
            for (var column = 0; column < columnCount; column++)
            {
                var abbreviation = abbreviations[weightRow, column];
                var weightFromExcel = weightsFromExcel[weightRow, column];
                var weight = weights[weightRow, column];

                if (weightFromExcel == null) continue;

                var absoluteColumn = startColumn + column;
                var columnLetter = absoluteColumn.GetColumnLetter();
                var addressLocation = RangeExtensions.GetAddressLocation(columnLetter, absoluteRow);

                if (double.IsNaN(weight))
                {
                    var message = $"Enter {BexConstants.StateProfileName.ToLower()} {weightLabel} in {addressLocation}" +
                                  $": <{weightFromExcel}> {BexConstants.NotRecognizedAsANumber}";
                    validation.AppendLine(message);
                    continue;
                }

                //Todo figure out how to send CW now that ID isn't nullable - sending zero is a problem
                var abbrevCode = abbreviation.ToString() == "CW"
                    ? new int()
                    : StateCodesFromBex.ReferenceData.Single(x => x.Abbreviation == abbreviation.ToString()).Id;

                var item = new StateDistributionItemPlus
                {
                    StateCode = abbrevCode,
                    Value = weight,
                    Location = addressLocation
                };
                Items.Add(item);
            }

            return validation;
        }

        public void WriteStatesIntoRange()
        {
            var content = GetStateContent();
            var rowCount = content.GetLength(0);

            var bodyRange = GetBodyRange();
            bodyRange.Offset[1, 0].Resize[bodyRange.Rows.Count - rowCount, bodyRange.Columns.Count].DeleteRangeUp();

            var inputLabelRange = GetInputLabelRange();
            inputLabelRange.Value2 = content;
        }

        public static string GetRangeName(int segmentId, int componentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.StateProfileRangeName}{componentId}";
        }

        public static string GetBasisRangeName(int segmentId, int componentId)
        {
            return
                $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.StateProfileRangeName}{componentId}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId)
        {
            return $"{GetRangeName(segmentId, componentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetSublinesRangeName(int segmentId, int componentId)
        {
            return GetSublinesRangeName(segmentId, componentId, ExcelConstants.StateProfileRangeName);
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId)
        {
            return GetSublinesHeaderRangeName(segmentId, componentId, ExcelConstants.StateProfileRangeName);
        }

        private static string[,] GetStateContent()
        {
            const string uslh = "USLH";
            var rowCount = StateCodesFromBex.ReferenceData.Count(x => x.Abbreviation != uslh);
            var content = new string[rowCount, 2];
            var counter = 0;
            StateCodesFromBex.ReferenceData.Where(x => x.Abbreviation != uslh).OrderBy(x => x.DisplayOrder).ForEach(x =>
            {
                content[counter, 0] = x.Abbreviation;
                content[counter, 1] = x.Name;
                counter++;
            });

            return content;
        }
    }
}