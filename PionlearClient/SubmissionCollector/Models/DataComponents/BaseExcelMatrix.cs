using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient.Extensions;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Models.DataComponents
{
    public abstract class BaseExcelMatrix : IExcelMatrix
    {
        public abstract string FriendlyName { get; }
        public virtual string FullName => FriendlyName;
        public virtual bool HasEmptyColumnToRight => true;
        public abstract string ExcelRangeName { get; }
        public bool IsEstimate { get; set; }
        public virtual string RangeName { get; set; }

        public virtual void Reformat()
        {
            var inputRange = GetInputRange();
            inputRange.SetToDefaultInputFont();
            inputRange.Locked = false;
            inputRange.ClearBorders();
            SetInputInteriorColorContemplatingEstimate(inputRange);
        }

        public abstract Range GetInputRange();
        public abstract Range GetInputLabelRange();
        public abstract StringBuilder Validate();

        public virtual void ToggleEstimate()
        {
            IsEstimate = !IsEstimate;

            var inputRange = GetInputRange();
            SetInputInteriorColorContemplatingEstimate(inputRange);
        }

        public void SetInputInteriorColorContemplatingEstimate(Range range)
        {
            if (IsEstimate)
            {
                range.SetInteriorColorWhenEstimate();
            }
            else
            {
                range.SetInputInteriorColor();
            }
        }

        public void SetInputDropdownInteriorColorContemplatingEstimate(Range range)
        {
            if (IsEstimate)
            {
                range.SetInteriorColorWhenEstimate();
            }
            else
            {
                range.SetInputDropdownInteriorColor();
            }
        }

        [JsonIgnore]
        public virtual bool HasData
        {
            get
            {
                var content = GetInputRange().GetContent();
                for (var row = 0; row < content.GetLength(0); row++)
                {
                    for (var column = 0; column < content.GetLength(1); column++)
                    {
                        if (content[row, column] != null) return true;
                    }
                }

                return false;
            }
        }

        [JsonIgnore] public int ColumnStart => RangeName.GetTopLeftCell().Column;

        public static string GetLineSublineConcatenatedNames(IList<ISubline> sublines)
        {
            var names = new List<string>();
            var lineNames = sublines.Select(subline => subline.LobShortName).Distinct();

            foreach (var lineName in lineNames)
            {
                var sublineNames = sublines.Where(subline => subline.LobShortName == lineName).Select(subline => subline.ShortName);
                var sublineNamesString = string.Join(", ", sublineNames);
                var lineNamesAndSublineNames = lineName.ConnectWithDash(sublineNamesString);
                names.Add(lineNamesAndSublineNames);
            }

            return string.Join(", ", names);
        }
    }
}