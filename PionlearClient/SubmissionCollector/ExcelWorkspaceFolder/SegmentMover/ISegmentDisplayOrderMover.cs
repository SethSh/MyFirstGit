using System.Collections.Generic;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ExcelWorkspaceFolder.SegmentMover
{
    internal interface ISegmentDisplayOrderMover
    {
        bool Validate(IList<ISegment> segments);
        void Move(IList<ISegment> segments, ISegment segment);
    }
}