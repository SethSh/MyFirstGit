using PionlearClient;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Models.Subline.ProfessionalLiability
{
    internal class SchoolBoardLegalSubline : BaseSubline
    {
        public SchoolBoardLegalSubline()
        {
            ImageSource = Resources.School.ToBitmapSource();
            Code = BexConstants.ProfessionalLiabilitySchoolBoardLegalSublineCode;
            ImageSource.Freeze();
        }
    }
}
