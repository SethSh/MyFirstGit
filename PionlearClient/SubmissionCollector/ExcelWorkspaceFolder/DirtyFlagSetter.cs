using System.Linq;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class DirtyFlagSetter
    {
        public void SetPackageIsDirty(Package package, bool isDirty)
        {
            package.IsDirty = isDirty;
            foreach (var x in package.Segments)
            {
                SetSegmentIsDirty(x, isDirty);
            }
        }

        private static void SetSegmentIsDirty(ISegment segment, bool isDirty)
        {
            segment.IsDirty = isDirty;
            foreach (var x in segment.ExcelComponents)
            {
                x.IsDirty = isDirty;
            }

            foreach (var providesLedger in segment.ExcelComponents.OfType<IProvidesLedger>())
            {
                providesLedger.Ledger.SetIsDirty(isDirty);
            }
        }
    }
}
