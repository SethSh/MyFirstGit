using System.Collections.Generic;

namespace SubmissionCollector.Extensions
{
    public static class ListExtensions
    {
        public static void RemoveAllButFirst<T>(this IList<T> list)
        {
            for (var row = list.Count - 1; row >= 0; row--)
            {
                if (row != 0) list.RemoveAt(row);
            }
        }

        public static void Swap<T>(this IList<T> list, int indexA, int indexB)
        {
            T tmp = list[indexA];
            list[indexA] = list[indexB];
            list[indexB] = tmp;
        }
    }
}
