using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using SubmissionCollector.Models.Comparers;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Extensions
{
    public static class ComparisonExtensions
    {
        public static bool IsNotEqualTo(this IEnumerable<string> first, IEnumerable<string> second)
        {
            return !first.IsEqualTo(second);
        }
        public static bool IsEqualTo(this IEnumerable<string> first, IEnumerable<string> second)
        {
            var firstList = first.ToList();
            var secondList = second.ToList();
            var list1 = new List<string>(firstList.Except(secondList));
            var list2 = new List<string>(secondList.Except(firstList));

            return !(list1.Any() || list2.Any());
        }

        public static bool IsNotEqualTo(this IEnumerable<ISubline> first, IEnumerable<ISubline> second)
        {
            return !first.IsEqualTo(second);
        }

        public static bool IsEqualTo(this IEnumerable<ISubline> first, IEnumerable<ISubline> second)
        {
            var firstList = first.ToList();
            var secondList = second.ToList();
            var list1 = new List<ISubline>(firstList.Except(secondList, new SublineComparer()));
            var list2 = new List<ISubline>(secondList.Except(firstList, new SublineComparer()));

            return !(list1.Any() || list2.Any());
        }

        public static bool IsEqualTo(this IEnumerable<UmbrellaTypeViewModel> first, IEnumerable<UmbrellaTypeViewModel> second)
        {
            var firstList = first.ToList();
            var secondList = second.ToList();
            var list1 = new List<UmbrellaTypeViewModel>(firstList.Except(secondList, new UmbrellaTypeComparer()));
            var list2 = new List<UmbrellaTypeViewModel>(secondList.Except(firstList, new UmbrellaTypeComparer()));

            return !(list1.Any() || list2.Any());
        }


    }
}
