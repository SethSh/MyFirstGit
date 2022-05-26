using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.KeyDataFolder;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector
{
    public class UserPreferences
    {
        public static string Filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexConstants.UserPreferencesFileName);

        public int InsertRowCountDefault = ExcelConstants.RowCountDefault;
        private const int PolicyProfileRowCountDefault = ExcelConstants.RowCountDefault;
        private const int TotalInsuredValueProfileRowCountDefault = ExcelConstants.RowCountDefault;
        public const int HistoricalPeriodCountDefault = ExcelConstants.RowCountDefault;
        public const int IndividualLossCountDefault = ExcelConstants.RowCountDefault;
        public const int RateChangeCountDefault = ExcelConstants.RowCountDefault;
        public const int WorkersCompClassCodeCountDefault = ExcelConstants.RowCountDefault;
        public const int WorkersCompStateAttachmentCountDefault = ExcelConstants.ColumnCountDefault;
        
        private const EvaluationDateType EvaluationDateTypeDefault = EvaluationDateType.None;
        private const UnderwritingYearType UnderwritingYearTypeDefault = UnderwritingYearType.None;
        private const int SubjectPolicyAlaeTreatmentDefault = BexConstants.LossOnlySubjectPolicyAlaeTreatmentId;
        private const bool SubmissionRibbonSelectionChoiceDefault = true;
        private const bool AllowSublineFlexibilityDefault = false;
        private const bool ShowMyCedentsDefault = false;
        private const bool ShowMyUnderwritersDefault = false;
        private const bool UseMeAsUnderwriterDefault = false;

        internal static bool FileExists => File.Exists(Filename); 
        internal static bool IsAggregateLossAndAlaeCombinedDefault = true;
        internal static bool IsAggregatePaidAvailableDefault = false;
        internal static bool IsIndividualLossAndAlaeCombinedDefault = false;
        internal static bool IsIndividualPaidAvailableDefault = false;
        internal static bool IsPolicyLimitAvailableDefault = true;
        internal static bool IsPolicyAttachmentAvailableDefault = true;
        internal static bool IsAccidentDateAvailableDefault = true;
        internal static bool IsPolicyDateAvailableDefault = true;
        internal static bool IsReportDateAvailableDefault = true;
        internal static bool IsEventCodeAvailableDefault = true;
        internal static bool ArePropertyNodesExpandedDefault = false;
        internal static bool CheckSynchronizationDefault = true;
        internal static bool IsAdminDefault = false;
        internal static double PaneWidthFactorDefault = 0.25;
        internal static PanePositionType PanePositionDefault = PanePositionType.Left;
        internal static short SegmentWorkSheetZoomDefault = 100;
        internal static bool IsTotalInsuredValueProfileExpandedDefault = false;

        public int ExposureBasis { get; set; }
        public int PeriodType { get; set; }
        public int ProfileBasisId { get; set; }
        public int PolicyProfileRowCount { get; set; }

        // ReSharper disable once StringLiteralTypo
        [JsonProperty("HistoricaPeriodCount")] public int HistoricalPeriodCount { get; set; }
        public int IndividualLossCount { get; set; }
        public int RateChangeCount { get; set; }
        public int WorkersCompClassCodeCount { get; set; }
        public int WorkersCompStateAttachmentCount { get; set; }
        public bool WorkersCompClassCodesIsActive { get; set; }
        public bool WorkersCompStateByHazardGroupIsIndependent { get; set; }
        public EvaluationDateType EvaluationDateType { get; set; }
        public UnderwritingYearType UnderwritingYearType { get; set; }
        public int SubjectPolicyAlaeTreatment { get; set; }
        public IList<BusinessPartner> MyCedents { get; set; }
        public bool ShowMyCedents { get; set; }
        public IList<Underwriter> MyUnderwriters { get; set; }
        public bool ShowMyUnderwriters { get; set; }
        public bool SubmissionRibbonSelectionChoice { get; set; }
        public bool AllowSublineFlexibility { get; set; }
        public bool IsAggregateLossAndAlaeCombined { get; set; }
        public bool IsAggregatePaidAvailable { get; set; }
        public bool IsIndividualLossAndAlaeCombined { get; set; }
        public bool IsIndividualPaidAvailable { get; set; }
        public bool IsPolicyLimitAvailable { get; set; }
        public bool IsPolicyAttachmentAvailable { get; set; }
        public bool IsAccidentDateAvailable { get; set; }
        public bool IsPolicyDateAvailable { get; set; }
        public bool IsReportDateAvailable { get; set; }
        public bool IsEventCodeAvailable { get; set; }
        public bool IsTotalInsuredValueProfileExpanded { get; set; }
        public bool UseMeAsUnderwriter { get; set; }
        public bool ArePropertyNodesExpanded { get; set; }
        public bool CheckSynchronization { get; set; }
        public bool IsAdmin { get; set; }
        public int InsertRowCount { get; set; }
        public double PaneWidthFactor { get; set; }
        public PanePositionType PanePosition { get; set; }
        public short SegmentWorksheetZoom { get; set; }
        public int TotalInsuredValueProfileRowCount { get; set; }


        public void CreateNew()
        {
            ExposureBasis = ExposureBasisFromBex.DefaultCode;
            PeriodType = HistoricalPeriodTypesFromBex.DefaultCode;
            ProfileBasisId = ProfileBasisFromBex.DefaultCode;
            InsertRowCount = InsertRowCountDefault;
            PolicyProfileRowCount = PolicyProfileRowCountDefault;
            TotalInsuredValueProfileRowCount = TotalInsuredValueProfileRowCountDefault;
            EvaluationDateType = EvaluationDateTypeDefault;
            UnderwritingYearType = UnderwritingYearTypeDefault;
            HistoricalPeriodCount = HistoricalPeriodCountDefault;
            IndividualLossCount = IndividualLossCountDefault;
            RateChangeCount = RateChangeCountDefault;
            WorkersCompClassCodeCount = WorkersCompClassCodeCountDefault;
            WorkersCompStateAttachmentCount = WorkersCompStateAttachmentCountDefault;
            WorkersCompClassCodesIsActive = ExcelConstants.WorkersCompUseClassCodesIsActiveDefault;
            WorkersCompStateByHazardGroupIsIndependent = ExcelConstants.WorkersCompStateByHazardGroupIsIndependentDefault;
            SubjectPolicyAlaeTreatment = SubjectPolicyAlaeTreatmentDefault;
            ShowMyCedents = ShowMyCedentsDefault;
            ShowMyUnderwriters = ShowMyUnderwritersDefault;
            SubmissionRibbonSelectionChoice = SubmissionRibbonSelectionChoiceDefault;
            AllowSublineFlexibility = AllowSublineFlexibilityDefault;
            IsAggregateLossAndAlaeCombined = IsAggregateLossAndAlaeCombinedDefault;
            IsAggregatePaidAvailable = IsAggregatePaidAvailableDefault;
            IsIndividualLossAndAlaeCombined = IsIndividualLossAndAlaeCombinedDefault;
            IsIndividualPaidAvailable = IsIndividualPaidAvailableDefault;
            IsPolicyLimitAvailable = IsPolicyLimitAvailableDefault;
            IsPolicyAttachmentAvailable = IsPolicyAttachmentAvailableDefault;
            IsAccidentDateAvailable = IsAccidentDateAvailableDefault;
            IsPolicyDateAvailable = IsPolicyDateAvailableDefault;
            IsReportDateAvailable = IsReportDateAvailableDefault;
            IsEventCodeAvailable = IsEventCodeAvailableDefault;
            IsTotalInsuredValueProfileExpanded = IsTotalInsuredValueProfileExpandedDefault;
            UseMeAsUnderwriter = UseMeAsUnderwriterDefault;
            ArePropertyNodesExpanded = ArePropertyNodesExpandedDefault;
            CheckSynchronization = CheckSynchronizationDefault;
            IsAdmin = IsAdminDefault;
            PaneWidthFactor = PaneWidthFactorDefault;
            PanePosition = PanePositionDefault;
            MyCedents = new List<BusinessPartner>();
            MyUnderwriters = new List<Underwriter>();
            SegmentWorksheetZoom = SegmentWorkSheetZoomDefault;
        }

        


        public void WriteToFile()
        {
            var json = JsonConvert.SerializeObject(this, Formatting.Indented);
            var fileInfo = new FileInfo(Filename);
            if (fileInfo.Directory == null) return;

            fileInfo.Directory.Create();
            File.WriteAllText(Filename, json);
        }

        public static UserPreferences ReadFromFile()
        {
            var json = File.ReadAllText(Filename);
            var userPreferences = JsonConvert.DeserializeObject<UserPreferences>(json);
            HandleNewProperties(userPreferences);
            return userPreferences;
        }

        private static void HandleNewProperties(UserPreferences userPreferences)
        {
            //for the properties that didn't always exist in user preferences
            const double paneWidthFactorWhenDoesNotExist = 0d;
            if (userPreferences.PaneWidthFactor.IsEqual(paneWidthFactorWhenDoesNotExist))
            {
                userPreferences.PaneWidthFactor = PaneWidthFactorDefault;
                userPreferences.PanePosition = PanePositionDefault;
            }

            if (userPreferences.RateChangeCount == 0)
            {
                userPreferences.RateChangeCount = RateChangeCountDefault;
            }

            if (userPreferences.WorkersCompClassCodeCount == 0)
            {
                userPreferences.WorkersCompClassCodeCount = WorkersCompClassCodeCountDefault;
            }

            if (userPreferences.WorkersCompStateAttachmentCount == 0)
            {
                userPreferences.WorkersCompStateAttachmentCount = WorkersCompStateAttachmentCountDefault;
            }

            if (userPreferences.SegmentWorksheetZoom == 0)
            {
                userPreferences.SegmentWorksheetZoom = SegmentWorkSheetZoomDefault;
            }

            if (userPreferences.TotalInsuredValueProfileRowCount == 0)
            {
                userPreferences.TotalInsuredValueProfileRowCount = TotalInsuredValueProfileRowCountDefault;
            }
        }
    }


    public enum PanePositionType
    {
        Left,
        Right
    }
}

