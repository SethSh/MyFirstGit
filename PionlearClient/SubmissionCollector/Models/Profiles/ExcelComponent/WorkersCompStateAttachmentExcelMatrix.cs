using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.CollectorClientPlus;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Profiles.ExcelComponent
{
    [JsonObject]
    public sealed class WorkersCompStateAttachmentExcelMatrix : SingleOccurrenceProfileExcelMatrix, IRangeResizable
    {
        public WorkersCompStateAttachmentExcelMatrix(int segmentId) : base(segmentId)
        {
            InterDisplayOrder = 39;
        }

        public override string FriendlyName => BexConstants.WorkersCompStateAttachmentProfileName;
        public override string ExcelRangeName => ExcelConstants.WorkersCompStateAttachmentProfileRangeName;

        
        [JsonIgnore]
        public IList<WorkersCompStateAttachmentAndWeightPlus> Items { get; set; }

        
        public override void Reformat()
        {
            base.Reformat();

            var headerRange = GetHeaderRange();
            headerRange.SetHeaderFormat();
            headerRange.ClearContents();
            headerRange.GetTopLeftCell().Value = BexConstants.WorkersCompStateAttachmentProfileName;

            var basisRange = GetProfileBasisRange();
            basisRange.Locked = true;
            basisRange.SetInputLabelColor();
            basisRange.Value = ProfileBasisFromBex.ReferenceData.Single(data => data.Id != BexConstants.PremiumProfileBasisId).Name;
            
            var bodyRange = GetBodyRange();
            bodyRange.ClearBorders();
            bodyRange.SetBorderToOrdinary();

            var bodyHeaderRange = GetBodyHeaderRange();
            bodyHeaderRange.ClearBorders();
            bodyHeaderRange.GetTopLeftCell().Locked = true;
            bodyHeaderRange.GetTopLeftCell().SetInputLabelColor();
            bodyHeaderRange.SetBorderToOrdinary();

            bodyHeaderRange.AlignLeft();
            bodyHeaderRange.GetTopLeftCell().Value = "Abbrev";
            bodyHeaderRange.GetTopLeftCell().Offset[0, 1].Value = "Name";
            bodyHeaderRange.GetTopLeftCell().Offset[0, 2].Value = "Primary";
            bodyHeaderRange.GetTopLeftCell().Offset[0, 2].AlignRight();
            bodyHeaderRange.Locked = true;
            bodyHeaderRange.SetInputLabelColor();
            bodyHeaderRange.GetRangeSubset(0, 3).AlignRight();
            bodyHeaderRange.GetRangeSubset(0, 3).SetInputColor();
            bodyHeaderRange.GetRangeSubset(0, 3).Locked = false;
            bodyHeaderRange.GetRangeSubset(0, 3).NumberFormat = FormatExtensions.WholeNumberFormat;

            var labelRange = GetInputLabelRange();
            labelRange.Locked = true;
            labelRange.SetInputLabelColor();
            labelRange.AlignLeft();
            labelRange.GetRangeSubset(0, 2).AlignRight();
            labelRange.GetRangeSubset(0, 2).NumberFormat = FormatExtensions.PercentFormatWithoutDecimal;
            
            var inputRange = GetInputRange();
            inputRange.AlignRight();
            inputRange.SetInputColor();
            inputRange.NumberFormat = FormatExtensions.PercentFormatWithoutDecimal;
            inputRange.SetBorderRightToResizable();
            
            labelRange.GetRangeSubset(0, 2).Formula = $"=1 - Sum({inputRange.GetFirstRow().Address[RowAbsolute:false, ColumnAbsolute:false]})";

            bodyRange.ColumnWidth = ExcelConstants.StandardColumnWidth;
            bodyRange.GetLastColumn().Offset[0, 1].ColumnWidth = ExcelConstants.MarginColumnWidth;
        }

        public override void ImplementProfileBasis()
        {
            
        }

        public Range GetInputAttachmentRange()
        {
            return GetInputRange().GetFirstRow().Offset[-1, 0];
        }

        public override Range GetInputRange()
        {
            return RangeName.GetRangeSubset(1, 3);
        }

        public override Range GetInputLabelRange()
        {
            var body = GetBodyRange();
            return body.Resize[body.Rows.Count, 3];
        }

        public override Range GetBodyHeaderRange()
        {
            return RangeName.GetRange().GetRow(0);
        }
        
        public override StringBuilder Validate()
        {
            var validation = new StringBuilder();
            Items = new List<WorkersCompStateAttachmentAndWeightPlus>();

            var gridRange = GetInputRange();
            var gridFromExcel = gridRange.GetContent();
            var gridAsDouble = gridFromExcel.ForceContentToDoubles();
            
            var attachmentLabelRange = GetBodyHeaderRange().GetRangeSubset(0,3);
            var attachmentsFromExcel = attachmentLabelRange.GetContent().GetRow(0).ToList();
            var attachmentsAsDouble = attachmentLabelRange.GetContent().ForceContentToDoubles().GetRow(0).ToList();

            var isAttachmentsAllNull = attachmentsFromExcel.All(att => att == null);
            var isGridAllNull = gridFromExcel.AllNull();
            if (isAttachmentsAllNull && isGridAllNull) return validation;

           
            var topLeftCell = attachmentLabelRange.GetTopLeftCell();
            var attachmentRow = topLeftCell.Row;
            var startColumn = topLeftCell.Column;
            
            topLeftCell = gridRange.GetTopLeftCell();
            var startRow = topLeftCell.Row;
            
            var rowCount = gridAsDouble.GetLength(0);
            var columnCount = gridAsDouble.GetLength(1);
            var suppressAttachmentValidation = new List<int>();
            var stateIds = GetStateIds();
            var columnLetters = GetColumnLetters(attachmentsAsDouble, startColumn);

            for (var row = 0; row < rowCount; row++)
            {
                var rowNumber = row + startRow;
                var stateId = stateIds[row];
                
                for (var column = 0; column < columnCount; column++)
                {
                    var attachment = attachmentsAsDouble[column];
                    var attachmentFromExcel = attachmentsFromExcel[column];
                    var gridItem = gridAsDouble[row, column];
                    var gridItemFromExcel = gridFromExcel[row, column];

                    if (gridItemFromExcel == null) continue;

                    var addressLocation = RangeExtensions.GetAddressLocation(columnLetters[column], rowNumber);
                    if (!double.IsNaN(gridItem))
                    {
                        //since looping thru rows, need to keep track of attachment validations not to repeat them
                        if (!suppressAttachmentValidation.Contains(column) && double.IsNaN(attachment))
                        {
                            var attachmentAddressLocation = RangeExtensions.GetAddressLocation(columnLetters[column], attachmentRow);
                            validation.AppendLine($"Enter {BexConstants.AttachmentName.ToLower()} in {attachmentAddressLocation}:" +
                                                  $" <{attachmentFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                            if (!suppressAttachmentValidation.Contains(column)) suppressAttachmentValidation.Add(column);
                        }
                    }
                    else
                    {
                        validation.AppendLine($"Enter {BexConstants.AttachmentName.ToLower()} value in {addressLocation}:" +
                                              $" <{gridItemFromExcel}> {BexConstants.NotRecognizedAsANumber}");
                    }
                    
                    if (gridItem.IsEqual(0)) continue;

                    Items.Add(new WorkersCompStateAttachmentAndWeightPlus
                    {
                        WorkersCompStateId = stateId,
                        Attachment = attachment,
                        Weight = gridItem,
                        Location = addressLocation
                    });
                }
            }
            
            return validation;
        }

        

        public static string GetRangeName(int segmentId)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{ExcelConstants.WorkersCompStateAttachmentProfileRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId)
        {
            return $"{GetRangeName(segmentId)}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetBasisRangeName(int segmentId)
        {
            return $"{GetRangeName(segmentId)}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public override void ToggleEstimate()
        {
            IsEstimate = !IsEstimate;

            var attachmentRange = GetInputAttachmentRange();
            SetInputInteriorColorContemplatingEstimate(attachmentRange);

            var inputRange = GetInputRange();
            SetInputInteriorColorContemplatingEstimate(inputRange);
        }

        public override Range GetBodyRange()
        {
            return RangeName.GetRangeSubset(1, 0);
        }

        private static Dictionary<int, string> GetColumnLetters(List<double> attachmentsAsDouble, int startColumn)
        {
            var columnLetters = new Dictionary<int, string>();
            for (var column = 0; column < attachmentsAsDouble.Count; column++)
            {
                columnLetters.Add(column, (startColumn + column).GetColumnLetter());
            }

            return columnLetters;
        }


        private List<int> GetStateIds()
        {
            var stateLabelRange = GetInputLabelRange();
            var stateAbbreviations = stateLabelRange.GetContent().ForceContentToStrings().GetColumn(0).ToList();
            var states = StateCodesFromBex.GetWorkersCompStates().ToDictionary(key => key.Abbreviation);
            var stateIds = new List<int>();
            foreach (var stateAbbreviation in stateAbbreviations)
            {
                stateIds.Add(states[stateAbbreviation].Id);
            }

            return stateIds;
        }
    }
}
