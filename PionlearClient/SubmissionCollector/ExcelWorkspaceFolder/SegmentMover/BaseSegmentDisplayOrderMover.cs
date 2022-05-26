using System.Collections.Generic;
using System.Linq;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder.SegmentMover
{
    internal abstract class BaseSegmentDisplayOrderMover: ISegmentDisplayOrderMover
    {
        public virtual bool Validate(IList<ISegment> segments)
        {
            var segment = segments.SingleOrDefault(s => s.IsSelected);
            if (segment != null) return true;

            MessageHelper.Show($"{BexConstants.SegmentName} must be selected in inventory tree", MessageType.Stop);
            return false;
        }

        public abstract void Move(IList<ISegment> segments, ISegment segment);
    }
}