using System.Collections.Generic;
using System.Linq;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder.SegmentMover
{
    internal class IncreaseSegmentDisplayOrder : BaseSegmentDisplayOrderMover
    {
        //e.g. move from 4 to 5
        public override bool Validate(IList<ISegment> segments)
        {
            var validation = base.Validate(segments);
            if (!validation) return false;
            
            var segment = segments.Single(s => s.IsSelected);

            var displayOrder = segment.DisplayOrder;
            if (displayOrder != segments.Min(x => x.DisplayOrder)) return true;

            MessageHelper.Show("Submission segment already at top", MessageType.Stop);
            return false;

        }

        public override void Move(IList<ISegment> segments, ISegment segment)
        {
            var displayOrder = segment.DisplayOrder;
            var otherSegment = segments.Single(x => x.DisplayOrder == displayOrder - 1);
            segments.Swap(displayOrder, displayOrder - 1);

            using (new WorkbookUnprotector())
            {
                segment.WorksheetManager.Worksheet.Move(Before: otherSegment.WorksheetManager.Worksheet);
            }
        }
    }
}
