using System;
using System.Collections.Generic;
using System.Linq;

namespace SubmissionCollector.Models.DataComponents
{
    internal static class ExcelComponentsExtensions
    {
        internal static bool ContainsGuid(this IEnumerable<IExcelComponent> excelComponents, Guid guid)
        {
            return excelComponents?.SingleOrDefault(y => y.Guid == guid) != null;
        }

        internal static IExcelComponent Get(this IEnumerable<IExcelComponent> excelComponents, Guid guid)
        {
            return excelComponents.SingleOrDefault(y => y.Guid == guid);
        }
    }
}
