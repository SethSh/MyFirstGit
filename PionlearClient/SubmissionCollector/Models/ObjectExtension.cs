namespace SubmissionCollector.Models
{
    internal static class ObjectExtension
    {
        internal static T[] GetRow<T>(this T[,] matrix, int rowIndex)
        {
            var columnCount = matrix.GetLength(1);
            var vector = new T[columnCount];

            for (var column = 0; column < columnCount; column++)
            {
                vector[column] = matrix[rowIndex, column];
            }
            return vector;
        }

        internal static T[] GetColumn<T>(this T[,] matrix, int columnIndex)
        {
            var rowCount = matrix.GetLength(0);
            var vector = new T[rowCount];

            for (var row = 0; row < rowCount; row++)
            {
                vector[row] = matrix[row, columnIndex];
            }
            return vector;
        }

        internal static object[,] GetColumns(this object[,] matrix, int[] columnIndices)
        {
            var rowCount = matrix.GetLength(0);
            var subset = new object[rowCount, columnIndices.Length];
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnIndices.Length; column++)
                {
                    var columnIndex = columnIndices[column];
                    subset[row, column] = matrix[row, columnIndex];
                }
            }
            return subset;
        }

        internal static bool AllNull<T>(this T[,] items)
        {
            var rowCount = items.GetLength(0);
            var columnCount = items.GetLength(1);
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    if (items[row, column] != null) return false;
                }
            }

            return true;
        }
    }
}
