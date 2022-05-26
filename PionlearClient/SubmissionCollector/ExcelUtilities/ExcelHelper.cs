using System;
using Microsoft.Office.Interop.Excel;

namespace SubmissionCollector.ExcelUtilities
{
    public static class ExcelHelper
    {
        private const double DotNetEquivalentToExcelError = double.NaN; 
        private const double DotNetEquivalentToExcelEmpty = double.NaN;
        private const double DotNetEquivalentToExcelNonInfinityString = double.NaN;
        
        public static object[,] ChangeToBaseZero(this object[,] baseOne)
        {
            var rowCount = baseOne.GetLength(0);
            var columnCount = baseOne.GetLength(1);
            var baseZero = new object[rowCount, columnCount];
            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < columnCount; j++)
                {
                    baseZero[i, j] = baseOne[i + 1, j + 1];
                }
            }
            return baseZero;
        }

        public static string ForceContentToString(this object obj)
        {
            return obj?.ToString() ?? string.Empty;
        }

        public static double[,] ForceContentToDoubles(this object[,] obj)
        {
            var rowCount = obj.GetLength(0);
            var columnCount = obj.GetLength(1);

            var data = new double[rowCount, columnCount];

            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < columnCount; j++)
                {
                    var rawDataItem = obj[i, j];
                    data[i, j] = ExcelElementToDouble(rawDataItem);
                }
            }

            return data;
        }

        public static double?[,] ForceContentToNullableDoubles(this object[,] obj)
        {
            var rowCount = obj.GetLength(0);
            var columnCount = obj.GetLength(1);

            var data = new double?[rowCount, columnCount];

            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < columnCount; j++)
                {
                    var rawDataItem = obj[i, j];
                    data[i, j] = ExcelElementToNullableDouble(rawDataItem);
                }
            }

            return data;
        }

        public static double?[] ForceContentToNullableDoubles(this object[] obj)
        {
            var data = new double?[obj.Length];

            for (var i = 0; i < obj.Length; i++)
            {
                var rawDataItem = obj[i];
                data[i] = ExcelElementToNullableDouble(rawDataItem);
            }

            return data;
        }

        public static DateTime?[,] ForceContentToNullableDates(this object[,] obj)
        {
            var rowCount = obj.GetLength(0);
            var columnCount = obj.GetLength(1);

            var data = new DateTime?[rowCount, columnCount];

            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < columnCount; j++)
                {
                    var rawDataItem = obj[i, j];
                    data[i, j] = ExcelElementToNullableDate(rawDataItem);
                }
            }

            return data;
        }

        public static string[,] ForceContentToStrings(this object[,] obj)
        {
            var rowCount = obj.GetLength(0);
            var columnCount = obj.GetLength(1);

            var strings = new string[rowCount, columnCount];

            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < columnCount; j++)
                {
                    var rawDataItem = obj[i, j];
                    if (rawDataItem != null) strings[i, j] = rawDataItem.ToString();
                }
            }

            return strings;
        }

        public static string[] ForceContentToStrings(this object[] obj)
        {
            var strings = new string[obj.Length];

            for (var i = 0; i < obj.Length; i++)
            {
                var rawDataItem = obj[i];
                if (rawDataItem != null) strings[i] = rawDataItem.ToString();
            }

            return strings;
        }

        public static object[,] GetContent(this Range range)
        {
            object[,] rangeContent;
            
            var rowCount = range.Rows.Count;
            var columnCount = range.Columns.Count;

            if (rowCount == 1 && columnCount == 1)
            {
                rangeContent = new object[1, 1];
                rangeContent[0, 0] = range.Value2;
            }
            else
            {
                var obj = (object[,])range.Value2;
                rangeContent = obj.ChangeToBaseZero();
            }
            return rangeContent;
        }

        private static double ExcelElementToDouble(object rawDataItem)
        {
            if (rawDataItem == null)
            {
                return DotNetEquivalentToExcelEmpty;
            }

            switch (Type.GetTypeCode(rawDataItem.GetType()))
            {
                case TypeCode.String:
                    return DotNetEquivalentToExcelString(rawDataItem.ToString());
                case TypeCode.Boolean:
                    return Convert.ToBoolean(rawDataItem) ? 1.0 : 0.0;
                case TypeCode.Int32:
                    return DotNetEquivalentToExcelError;
                default:
                    return Convert.ToDouble(rawDataItem.ToString());
            }
        }

        private static double? ExcelElementToNullableDouble(object rawDataItem)
        {
            if (rawDataItem == null)
            {
                return null;
            }

            switch (Type.GetTypeCode(rawDataItem.GetType()))
            {
                case TypeCode.String:
                    return DotNetEquivalentToExcelString(rawDataItem.ToString());
                case TypeCode.Boolean:
                    return Convert.ToBoolean(rawDataItem) ? 1.0 : 0.0;
                case TypeCode.Int32:
                    return DotNetEquivalentToExcelError;
                default:
                    return Convert.ToDouble(rawDataItem.ToString());
            }
        }

        private static DateTime? ExcelElementToNullableDate(object rawDataItem)
        {
            if (rawDataItem == null)
            {
                return null;
            }

            try
            {
                switch (Type.GetTypeCode(rawDataItem.GetType()))
                {
                    case TypeCode.String:
                        return null;
                    case TypeCode.Boolean:
                        return null;
                    case TypeCode.Int32:
                        return null;
                    default:
                        return DateTime.FromOADate(Convert.ToDouble(rawDataItem));
                }
            }
            catch
            {
                return null;
            }
        }

        private static double DotNetEquivalentToExcelString(string excelString)
        {
            switch (excelString.ToLower())
            {
                case "inf":
                    return double.PositiveInfinity;
                case "-inf":
                    return double.NegativeInfinity;
            }
            return DotNetEquivalentToExcelNonInfinityString;
        }
    }
}
