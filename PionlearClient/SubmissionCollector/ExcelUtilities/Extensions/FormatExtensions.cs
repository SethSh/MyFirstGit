using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;

namespace SubmissionCollector.ExcelUtilities.Extensions
{
    public static class FormatExtensions
    {
        public const string StringFormat = "General";
        public const string PercentFormat = "0.00%";
        public const string PercentFormatWithoutDecimal = "0%";
        public const string DateFormat = "m/d/yyyy";
        public const string YearFormat = "0000";
        public const string WholeNumberFormat = "#,##0";
        public const string WholeNumberWithFourLeadingZerosFormat = "0000";
        public const string FactorFormat = "0.000";

        public const string ColumnWidthDateDefault = "15";
        public const string ColumnWidthFactorDefault = "15";
        public const string ColumnWidthEmptyDefault = "5";
        
        public static readonly Color LabelInteriorColor = Color.FromArgb(236, 236, 236);
        public static readonly Color SublineHeaderInteriorColor = Color.FromArgb(91, 155, 213);
        public static readonly Color SublineHeaderFontColor = Color.White;

        private static readonly Color InputInteriorColor = Color.FromArgb(255, 255, 204);
        private static readonly Color InputDropdownInteriorColor = Color.FromArgb(255, 204, 153);
        private static readonly Color EstimateInteriorColor = Color.BurlyWood;
        private static readonly Color InputFontColor = Color.Blue;
        private static readonly Color NonInputFontColor = Color.Black;
        private static readonly Color DisabledFontColor = Color.LightGray;

        private const string DefaultFontSize = "11";
        private const string DefaultFontName = "Calibri";

        public static void ClearConditionalFormats(this Range range)
        {
            range.FormatConditions.Delete();
        }
        
        public static void DeEmphasizeZero(this Range range)
        {
            range.ClearConditionalFormats();
            var fc = (FormatCondition)range.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, "=0");
            fc.Font.Color = DisabledFontColor;
        }

        public static void ClearBorders(this Range range)
        {
            foreach (Border border in range.Borders)
            {
                border.ClearBorder();
            }
        }

        public static void ClearInteriorColor(this Range range)
        {
            range.Interior.Pattern = XlPattern.xlPatternNone;
            range.Interior.TintAndShade = 0;
        }

        public static void ClearFontColor(this Range range)
        {
            range.Font.Color = NonInputFontColor;
        }

        public static void SetSublineHeaderFormat(this Range range)
        {
            range.SetSublineHeaderColor();
            range.Locked = true;
            range.SetFactorColumnWidth();
            range.SetBorderAroundToOrdinary();
            range.AlignCenterAcrossSelection();
            range.Font.Bold = true;
        }

        public static void SetHeaderFormat(this Range range)
        {
            range.SetBorderAroundToOrdinary();
            range.Locked = true;
            range.SetInputLabelColor(); 
            range.AlignCenterAcrossSelection();
            range.Font.Bold = true;
            range.SetToDefaultFont();
        }

        public static void SetBodyHeaderFormat(this Range range)
        {
            range.SetBorderAroundToOrdinary();
            range.Locked = true;
            range.SetInputLabelColor();
            range.SetToDefaultFont();
        }

        public static void SetSublineFormat(this Range range)
        {
            range.SetInputLabelColor();
            range.Locked = true;
            range.SetBorderAroundToOrdinary();
            range.AlignCenterAcrossSelection();
        }

        public static void SetInputLabelFormatWithBorder(this Range range)
        {
            range.SetBorderAroundToOrdinary();
            range.SetInputLabelFormat();
        }

        public static void SetInputLabelFormat(this Range range)
        {
            range.Locked = true;
            range.SetInputLabelColor();
        }

        public static void SetSublineHeaderColor(this Range range)
        {
            range.Interior.Color = SublineHeaderInteriorColor;
            range.Font.Color = SublineHeaderFontColor;
        }

        public static void SetInputLabelColor(this Range range)
        {
            range.SetInputLabelFontColor();
            range.SetInputLabelInteriorColor();
        }

        public static void SetInputLabelInteriorColor(this Range range)
        {
            range.Interior.Color = LabelInteriorColor;
        }

        public static void SetDateColumnWidth(this Range range)
        {
            range.ColumnWidth = ColumnWidthDateDefault;
        }

        public static void SetFactorColumnWidth(this Range range)
        {
            range.ColumnWidth = ColumnWidthFactorDefault;
        }

        public static void SetFactorColumnWidth(this Range range, int width)
        {
            range.ColumnWidth = width;
        }

        public static void SetRenewedInputColor(this Range range)
        {
            range.SetRenewedInputInteriorColor();
            range.SetRenewedInputFontColor();
        }

        private static void SetRenewedInputInteriorColor(this Range range)
        {
            range.Interior.Color = Color.Green;
        }

        public static void SetInputDropdownInteriorColor(this Range range)
        {
            range.Interior.Color = InputDropdownInteriorColor;
        }

        public static void SetInputLabelFontColor(this Range range)
        {
            range.Font.Color = Color.Black;
        }
        
        public static void SetInteriorColorWhenEstimate(this Range range)
        {
            range.Interior.Color = EstimateInteriorColor;
        }

        public static void SetInputInteriorColor(this Range range)
        {
            range.Interior.Color = InputInteriorColor;
        }

        public static void SetBorderAroundToResizable(this Range range)
        {
            range.SetBorderToOrdinary();
            range.SetBorderBottomToResizable();
        }

        public static void SetBorderToOrdinary(this Range range)
        {
            range.ClearBorders();
            range.BorderAround();
        }

        public static void SetBorderBottomToOrdinary(this Range range)
        {
            range.ClearBorder(XlBordersIndex.xlEdgeBottom);
            range.SetBorderToThin(XlBordersIndex.xlEdgeBottom);
        }

        public static void SetBorderTopToOrdinary(this Range range)
        {
            range.Borders[XlBordersIndex.xlEdgeTop].ClearBorder();
            range.SetBorderToThin(XlBordersIndex.xlEdgeTop);
        }

        public static void SetToDefaultFont(this Range range)
        {
            range.Font.Name = DefaultFontName;
            range.Font.Size = DefaultFontSize;
        }

        public static void SetToDefaultInputFont(this Range range)
        {
            range.SetToDefaultFont();
            range.Font.Size = DefaultFontSize;
            range.Font.Bold = false;
            range.SetInputFontColor();
        }

        public static void SetBorderAroundToOrdinary(this Range range)
        {
            ClearBorders(range);
            SetBorderLeftToOrdinary(range);
            SetBorderRightToOrdinary(range);
            SetBorderBottomToOrdinary(range);
            SetBorderTopToOrdinary(range);
        }

        public static void ClearBorderLeft(this Range range)
        {
            range.ClearBorder(XlBordersIndex.xlEdgeLeft);
        }

        public static void ClearBorderRight(this Range range)
        {
            range.ClearBorder(XlBordersIndex.xlEdgeRight);
        }

        public static void ClearBorderAllButTop(this Range range)
        {
            range.ClearBorder(XlBordersIndex.xlEdgeRight);
            range.ClearBorder(XlBordersIndex.xlEdgeLeft);
            range.ClearBorder(XlBordersIndex.xlEdgeBottom);
        }

        public static void SetBorderLeftToOrdinary(this Range range)
        {
            range.ClearBorderLeft();
            range.SetBorderToThin(XlBordersIndex.xlEdgeLeft);
        }

        public static void SetBorderRightToOrdinary(this Range range)
        {
            range.ClearBorderRight();
            range.SetBorderToThin(XlBordersIndex.xlEdgeRight);
        }

        public static void SetBorderRightToResizable(this Range range)
        {
            range.SetBorderToThin(XlBordersIndex.xlEdgeBottom);
            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDot;
        }
        
        public static void SetBorderBottomToResizable(this Range range)
        {
            range.SetBorderToThin(XlBordersIndex.xlEdgeBottom);
            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDot;
        }

        public static void SetInputFormat(this Range range)
        {
            range.SetInputColor();
            range.Locked = false;
        }

        public static void SetInputColor(this Range range)
        {
            range.SetInputFontColor();
            range.SetInputInteriorColor();
        }

        public static void SetInputFontColor(this Range range)
        {
            range.Font.Color = InputFontColor;
        }
        
        public static string[,] ToNByOneArray(this IList<string> list)
        {
            return list.ToArray().ToNByOneArray();
        }
        
        public static string[,] ToNByOneArray(this string[] s)
        {
            var length = s.Length;
            var s2 = new string[length, 1];
            
            for (var row = 0; row < length; row++)
            {
                s2[row, 0] = s[row];
            }
            return s2;
        }

        public static string[,] ToOneByNArray(this IList<string> list)
        {
            return list.ToArray().ToOneByNOneArray();
        }

        public static void FormatWithDates(this Range range)
        {
            range.NumberFormat = DateFormat;
        }

        public static void FormatWithWholeNumbers(this Range range)
        {
            range.NumberFormat = WholeNumberFormat;
        }
        public static void FormatWithPercents(this Range range)
        {
            range.NumberFormat = PercentFormat;
        }

        public static void FormatWithFactors(this Range range)
        {
            range.NumberFormat = FactorFormat;
        }

        public static void AlignRight(this Range range)
        {
            range.HorizontalAlignment = XlHAlign.xlHAlignRight;
        }

        public static void AlignLeft(this Range range)
        {
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        public static void AlignCenter(this Range range)
        {
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public static void AlignCenterAcrossSelection(this Range range)
        {
            range.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
        }

        public static void ClearSumCheck(this Range range)
        {
            range.FormatConditions.Delete();
        }

        public static void ApplySumCheck(this Range range)
        {
            var iconSetCondition = (IconSetCondition)range.FormatConditions.AddIconSetCondition();
            iconSetCondition.IconCriteria[1].Icon = XlIcon.xlIconRedCross;
            iconSetCondition.IconCriteria[2].Type = XlConditionValueTypes.xlConditionValueNumber;
            iconSetCondition.IconCriteria[2].Value = 1- NumericalConstants.ValidationProfileTolerance;
            iconSetCondition.IconCriteria[2].Operator = 7;
            iconSetCondition.IconCriteria[2].Icon = XlIcon.xlIconGreenCheck;

            iconSetCondition.IconCriteria[3].Type = XlConditionValueTypes.xlConditionValueNumber;
            iconSetCondition.IconCriteria[3].Value = 1+ NumericalConstants.ValidationProfileTolerance;
            iconSetCondition.IconCriteria[3].Operator = 5;
            iconSetCondition.IconCriteria[3].Icon = XlIcon.xlIconRedCross;
        }

        private static void SetRenewedInputFontColor(this Range range)
        {
            range.Font.Color = Color.Yellow;
        }

        private static string[,] ToOneByNOneArray(this IReadOnlyList<string> s)
        {
            var length = s.Count;
            var s2 = new string[1, length];

            for (var row = 0; row < length; row++)
            {
                s2[0, row] = s[row];
            }
            return s2;
        }

        private static void SetBorderToThin(this Range range, XlBordersIndex xlBordersIndex)
        {
            range.Borders[xlBordersIndex].Weight = XlBorderWeight.xlThin;
        }

        private static void ClearBorder(this Range range, XlBordersIndex xlBordersIndex)
        {
            range.Borders[xlBordersIndex].ClearBorder();
        }
        private static void ClearBorder(this Border border)
        {
            border.LineStyle = XlLineStyle.xlLineStyleNone;
        }
    }
}
