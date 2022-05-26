using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace SubmissionCollector.ExcelUtilities.Extensions
{
    public static class RangeExtensions
    {
        public const string DownArrow = "\u2193";
        public const string RightArrow = "\u2192";

        private static readonly Application ExcelApplication = Globals.ThisWorkbook.Application;

        internal static string GetAddressLocation(string columnLetter, int rowNumber)
        {
            return $"{BexConstants.RangeName.ToLower()} {columnLetter}{rowNumber}";
        }

        internal static string GetAddressLocation(IList<string> columnLetters, int rowNumber)
        {
            var sb = new StringBuilder();
            sb.Append($"{BexConstants.RangeName.ToLower()}");
            sb.Append(" ");
            sb.Append(string.Join(",", columnLetters.Select(columnLetter => $"{columnLetter}{rowNumber}").ToArray()));
            return sb.ToString();
        }

        public static void DeleteRangeUp(this Range range)
        {
            range.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        public static void InsertRangeDown(this Range range)
        {
            ClearTheClipboard(); 
            range.Insert(XlInsertShiftDirection.xlShiftDown);
        }


        public static void InsertColumnsToRight(this Range range)
        {
            ClearTheClipboard();
            range.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
        }

        private static void ClearTheClipboard()
        {
            //[STAThread] issue: Current thread must be set to single thread apartment (STA) mode before OLE calls can be made.
            var thread = new Thread(Clipboard.Clear);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        public static void DeleteRangeName(this string rangeName)
        {
            Globals.ThisWorkbook.Names.Item(rangeName).Delete();
        }

        public static void RenameRange(this string rangeName, string newRangeName)
        {
            rangeName.GetName().Name = newRangeName;
        }

        public static Range GetRange(this Worksheet ws, string address)
        {
            return ws.Range[address];
        }

        public static Range GetRange(this string rangeName)
        {
            return Globals.ThisWorkbook.Names.Item(rangeName).RefersToRange;
        }

        public static Range GetRangeOrDefault(this string rangeName)
        {
            return !rangeName.ExistsInWorkbook() ? null : Globals.ThisWorkbook.Names.Item(rangeName).RefersToRange;
        }

        public static Range GetRangeSubset(this string rangeName, int rowStart, int columnStart)
        {
            var range = rangeName.GetRange();
            var rowCount = range.Rows.Count;
            var columnCount = range.Columns.Count;

            return range.Offset[rowStart, columnStart].Resize[rowCount - rowStart, columnCount - columnStart];
        }

        public static object[,] GetContentFormulas(this Range range)
        {
            return ((object[,]) range.Formula).ChangeToBaseZero();
        }

        public static Range GetRangeSubset(this Range range, int rowStart, int columnStart)
        {
            var rowCount = range.Rows.Count;
            var columnCount = range.Columns.Count;

            return range.Offset[rowStart, columnStart].Resize[rowCount - rowStart, columnCount - columnStart];
        }

        public static Range GetTopRightCell(this string rangeName)
        {
            return rangeName.GetRange().GetTopRightCell();
        }

        public static Range GetTopRightCell(this Range range)
        {
            return range.Resize[1, 1].Offset[0, range.Columns.Count - 1];
        }

        public static int GetTopRow(this Range range)
        {
            return range.GetTopLeftCell().Row;
        }

        public static int GetBottomRow(this Range range)
        {
            return range.GetBottomLeftCell().Row;
        }

        public static Range GetTopLeftCell(this string rangeName)
        {
            return rangeName.GetRange().GetTopLeftCell();
        }

        public static Range GetColumn(this Range range, int i)
        {
            return range.GetTopLeftCell().Offset[0, i].Resize[range.Rows.Count, 1];
        }

        public static Range GetRow(this Range range, int i)
        {
            return range.GetTopLeftCell().Offset[i, 0].Resize[1, range.Columns.Count];
        }

        public static Range GetLastColumn(this Range range)
        {
            return range.GetTopRightCell().Resize[range.Rows.Count, 1];
        }

        public static Range GetTopLeftCell(this Range range)
        {
            return range.Resize[1, 1];
        }

        public static Range GetBottomRightCell(this Range range)
        {
            return range.Resize[1, 1].Offset[range.Rows.Count - 1, range.Columns.Count - 1];
        }

        public static Range GetBottomLeftCell(this Range range)
        {
            return range.Resize[1, 1].Offset[range.Rows.Count - 1, 0];
        }

        public static bool ContainsRange(this string rangeName, Range range)
        {
            var largeRange = rangeName.GetRange();
            return ExcelApplication.Intersect(range, largeRange) != null;
        }

        public static void ResetSourceRangeNames(this Worksheet worksheet, string from, string to)
        {
            if (worksheet == null)
            {
                const string message = "Can't find worksheet that contains range names";
                throw new ArgumentNullException(message);
            }

            foreach (Name item in worksheet.Names)
            {
                var rangeNameWithSheetName = item.Name;
                var prefixEndPosition = rangeNameWithSheetName.IndexOf("!", StringComparison.Ordinal);
                var rangeName = rangeNameWithSheetName.Substring(prefixEndPosition + 1);

                var newRangeName = rangeName.Replace(from, to);

                item.RefersToRange.SetInvisibleRangeName(newRangeName);
                Globals.ThisWorkbook.DeleteFromNameCollection(item);
            }
        }

        public static Name GetName(this string rangeName)
        {
            return ExcelApplication.Names.Item(rangeName);
        }

        public static bool IsTopRowEqual(this Range range1, Range range2)
        {
            return range1.GetTopLeftCell().Row == range2.GetTopLeftCell().Row;
        }

        public static bool DoesRangeIntersect(this Range range1, Range range2)
        {
            return ExcelApplication.Intersect(range1, range2) != null;
        }

        public static bool IsSelectionIntersectingLastColumn(this Range range)
        {
            var lastRow = range.GetLastColumn();
            var selectedRange = (Range) ExcelApplication.Selection;
            return ExcelApplication.Intersect(lastRow, selectedRange) != null;
        }

        public static Range GetFirstRow(this Range range)
        {
            return range.GetRow(0);
        }

        public static Range GetFirstRows(this Range range, int rowCount)
        {
            return range.GetTopLeftCell().Resize[rowCount, range.Columns.Count];
        }

        public static Range GetFirstColumn(this Range range)
        {
            return range.GetColumn(0);
        }

        public static Range GetFirstColumns(this Range range, int columnCount)
        {
            return range.Resize[range.Rows.Count, columnCount];
        }

        public static Range GetLastRow(this Range range)
        {
            return range.GetLastRows(1);
        }

        public static Range GetLastRows(this Range range, int count)
        {
            var offset = -(count - 1);
            return range.GetBottomLeftCell().Offset[offset, 0].Resize[count, range.Columns.Count];
        }

        public static void MoveRangeContent(this Range fromRange, Range toRange)
        {
            toRange.Value = fromRange.Value;
            fromRange.ClearContents();
        }

        public static Range AppendColumn(this Range range)
        {
            return range.Resize[range.Rows.Count, range.Columns.Count + 1];
        }

        public static Range AppendColumnsToLeft(this Range range, int additionalColumnCount)
        {
            return range.Offset[0, -additionalColumnCount].Resize[range.Rows.Count, range.Columns.Count + additionalColumnCount];
        }

        public static Range AppendRow(this Range range)
        {
            return range.Resize[range.Rows.Count + 1, range.Columns.Count];
        }

        public static Range RemoveLastRow(this Range range)
        {
            return range.RemoveLastRows(1);
        }

        public static Range RemoveLastRows(this Range range, int count)
        {
            return range.Resize[range.Rows.Count - count, range.Columns.Count];
        }

        public static Range RemoveLastColumn(this Range range)
        {
            return range.Resize[range.Rows.Count, range.Columns.Count - 1];
        }

        public static void ShowColumns(this Range range)
        {
            range.Resize[1, range.Columns.Count].EntireColumn.Hidden = false;
        }

        public static void HideColumns(this Range range)
        {
            range.Resize[1, range.Columns.Count].EntireColumn.Hidden = true;
        }

        public static Range Union(this Range range, Range otherRange)
        {
            return Globals.ThisWorkbook.Application.Union(range, otherRange);
        }

        public static void SetWorkbookRangeNamesVisibility(IList<string> filters, bool isVisible)
        {
            var wb = Globals.ThisWorkbook;
            var names = wb.Names.Cast<Name>().ToList();

            foreach (var filter in filters)
            {
                foreach (var name in names.Where(nm => nm.Name.Contains(filter)))
                {
                    if (name.Visible != isVisible) name.Visible = isVisible;
                }
            }
        }

        public static void SetInvisibleRangeName(this Range range, string rangeName)
        {
            range.Name = rangeName;

            var wb = Globals.ThisWorkbook;
            wb.Names.Item(rangeName).Visible = false;
        }

        public static bool ExistsInWorkbook(this string rangeName)
        {
            var rangeNames = Globals.ThisWorkbook.Names.Cast<Name>().Select(n => n.Name);
            return rangeNames.Contains(rangeName);
        }

        public static bool IntersectsNamedRange(this Range range, string rangeName)
        {
            var application = Globals.ThisWorkbook.Application;
            return application.Intersect(rangeName.GetName().RefersToRange, range) != null;
        }

        public static bool IntersectsAnyNamedRanges(this Range range, IEnumerable<string> rangeNames)
        {
            return range.GetIntersectingRangeNames(rangeNames).Any();
        }

        public static IEnumerable<string> GetIntersectingRangeNames(this Range range, IEnumerable<string> rangeNames)
        {
            return rangeNames.Where(range.IntersectsNamedRange);
        }

        public static IList<string> GetMatchingRangeNames(Regex regex)
        {
            var list = new List<string>();
            //here I don't need to be cognizant of name.Name.StartsWith("_xlfn.")
            foreach (var item in Globals.ThisWorkbook.Names.Cast<Name>())
            {
                if (regex.Match(item.Name).Success) list.Add(item.Name);
            }

            return list;
        }

        public static IList<string> GetMatchingRangeNames(Regex regex, IEnumerable<Name> names)
        {
            var list = new List<string>();
            //here I don't need to be cognizant of name.Name.StartsWith("_xlfn.")
            foreach (var item in names)
            {
                if (regex.Match(item.Name).Success) list.Add(item.Name);
            }

            return list;
        }

        public static string GetColumnLetter(this int column)
        {
            string letter = null;

            var columnIndex = column;
            do
            {
                var remainder = (columnIndex - 1) % 26;
                letter = (char) (remainder + 65) + letter;
                columnIndex = (columnIndex - remainder) / 26;
            } while (columnIndex > 0);

            return letter;
        }

        public static IDictionary<int, string> GetColumnLetters(this int columnStart, int columnCount)
        {
            var dict = new Dictionary<int, string>();

            for (var i = 0; i < columnCount; i++)
            {
                var columnIndex = columnStart + i;
                dict.Add(columnIndex, columnIndex.GetColumnLetter());
            }

            return dict;
        }


        public static void LockNecessaryCells(this Range range)
        {
            var labelColorInExcel = ColorTranslator.ToOle(FormatExtensions.LabelInteriorColor);
            var headerLabelColorInExcel = ColorTranslator.ToOle(FormatExtensions.SublineHeaderInteriorColor);

            foreach (Range r in range)
            {
                var color = r.Interior.Color;
                if (color.Equals(labelColorInExcel) || color.Equals(headerLabelColorInExcel))
                {
                    r.Locked = true;
                }
            }
        }

        public static bool IsEmpty(this Range range)
        {
            var content = range.GetContent().ForceContentToStrings();
            for (var row = 0; row < content.GetLength(0); row++)
            {
                for (var column = 0; column < content.GetLength(1); column++)
                {
                    if (!string.IsNullOrEmpty(content[row, column]))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public static bool IsNotAllZero(this Range range)
        {
            return !range.IsAllZero();
        }

        public static bool IsAllZero(this Range range)
        {
            var content = range.GetContent().ForceContentToDoubles();
            for (var row = 0; row < content.GetLength(0); row++)
            {
                for (var column = 0; column < content.GetLength(1); column++)
                {
                    var item = content[row, column];
                    if (!double.IsNaN(item) && !item.IsEqual(0))
                    {
                        return false;
                    }
                }
            }
            return true;
        }
    }
}
