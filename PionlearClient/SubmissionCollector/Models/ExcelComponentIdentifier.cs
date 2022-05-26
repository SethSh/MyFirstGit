using SubmissionCollector.Enums;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Package.DataComponents;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models
{
    internal class ExcelComponentIdentifier
    {
        public IExcelMatrix ExcelMatrix { get; set; }
        public bool Validate()
        {
            var packageIdentifier = new PackageExcelComponentIdentifier();
            if (packageIdentifier.Validate(true))
            {
                ExcelMatrix = packageIdentifier.ExcelMatrix;
                return true;
            }

            var segmentIdentifier = new SegmentExcelComponentIdentifier();
            if (segmentIdentifier.Validate(true))
            {
                ExcelMatrix = segmentIdentifier.ExcelMatrix;
                return true;
            }

            MessageHelper.Show("Selected cell not identified as within a data component", MessageType.Stop);
            return false;
        }
    }
}
