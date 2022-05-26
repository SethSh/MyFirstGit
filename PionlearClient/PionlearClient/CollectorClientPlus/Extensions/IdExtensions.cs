using System.Collections.Generic;
using System.Linq;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    internal static class IdExtensions
    {
        internal static bool IsEqualsTo(this ICollection<long> ids, ICollection<long> otherIds)
        {
            return ids.Count == otherIds.Count && otherIds.All(ids.Contains);
        }
    }
}
