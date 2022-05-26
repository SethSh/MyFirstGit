using System.Collections.Generic;
using PionlearClient.KeyDataFolder;
using Brushes = System.Windows.Media.Brushes;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignUserPreferencesViewModel : BaseUserPreferencesViewModel
    {
        public DesignUserPreferencesViewModel()
        {
            ExposureBasis = 1;
            PeriodType = 1;
            ProfileBasisId = 1;
            PolicyProfileCount = 10;
            PolicyProfileCountBackground = Brushes.Red;
            ReportDateBackground = Brushes.Red;
            EvaluationDateType = EvaluationDateType.EndOfLastYear;
            OkButtonEnabled = true;
            OkButtonMessage = "Message goes here";
            ShowMyCedents = true;
            SubmissionRibbonSelectionChoice = true;
            AllowSublineFlexibility = true;
            IsAggregateLossAndAlaeCombined = UserPreferences.IsAggregateLossAndAlaeCombinedDefault;
            IsAggregatePaidAvailable = UserPreferences.IsAggregatePaidAvailableDefault;
            IsIndividualLossAndAlaeCombined = UserPreferences.IsIndividualLossAndAlaeCombinedDefault;
            IsIndividualPaidAvailable = UserPreferences.IsIndividualPaidAvailableDefault;
            IsPolicyLimitAvailable = UserPreferences.IsPolicyLimitAvailableDefault;
            IsPolicyAttachmentAvailable = UserPreferences.IsPolicyAttachmentAvailableDefault;
            IsAccidentDateAvailable = UserPreferences.IsAccidentDateAvailableDefault;
            IsPolicyDateAvailable = UserPreferences.IsPolicyDateAvailableDefault;
            IsReportDateAvailable = UserPreferences.IsReportDateAvailableDefault;
            IsTotalInsuredValueProfileExpanded = UserPreferences.IsTotalInsuredValueProfileExpandedDefault;
            UseMeAsUnderwriter = false;
            ArePropertyNodesExpanded = UserPreferences.ArePropertyNodesExpandedDefault;
            MyCedents = new List<BusinessPartner>
            {
                new BusinessPartner {Id = "0000123456", Name = "Seth"},
                new BusinessPartner {Id = "0000abcdef", Name = "Jeff"}
            };
        }
    }
}
