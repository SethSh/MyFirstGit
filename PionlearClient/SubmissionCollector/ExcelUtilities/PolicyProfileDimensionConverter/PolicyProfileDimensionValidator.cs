using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter
{
    public class PolicyProfileDimensionValidator
    {
        public PolicyExcelMatrix PolicyExcelMatrix;
        public bool Validate()
        {
            var identifier = new SegmentExcelComponentIdentifier();
            if (!identifier.Validate()) return false;

            PolicyExcelMatrix = identifier.ExcelMatrix as PolicyExcelMatrix;
            if (PolicyExcelMatrix == null)
            {
                MessageHelper.Show($"The selection must be within a {BexConstants.PolicyProfileName.ToLower()} range", MessageType.Stop);
                return false;
            }
            return true;
        }
    }
}
