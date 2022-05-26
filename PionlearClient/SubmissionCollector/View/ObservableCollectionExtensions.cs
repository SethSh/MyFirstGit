using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace SubmissionCollector.View
{
    internal static class ObservableCollectionExtensions
    {
        public static void Sort<TSource, TKey>(this ObservableCollection<TSource> source, Func<TSource, TKey> keySelector)
        {
            var sortedList = source.OrderBy(keySelector).ToList();
            source.Clear();
            foreach (var sortedItem in sortedList)
            {
                source.Add(sortedItem);
            }
        }
    }
}
