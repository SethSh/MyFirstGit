using System.Collections.Generic;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Models.Comparers
{
    internal class SublineComparer : IEqualityComparer<ISubline>
    {
        public bool Equals(ISubline x, ISubline y)
        {
            return y != null && x != null && x.Code == y.Code;
        }

        public int GetHashCode(ISubline obj)
        {
            return obj.Code.GetHashCode();
        }
    }
}