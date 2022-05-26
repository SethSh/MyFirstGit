namespace SubmissionCollector
{
    internal static class ExcelConstants
    {
        public const int WorksheetNameMaximumCharacters = 30;
        public const string SegmentTemplateWorksheetName = "SegmentTemplate";
        public const string SegmentTemplateRangeName = "segmentTemplate";

        public const int RowCountDefault = 10;
        public const int ColumnCountDefault = 6;
        public const int InBetweenRowCount = 2;
        public const bool WorkersCompUseClassCodesIsActiveDefault = false;
        public const bool WorkersCompStateByHazardGroupIsIndependentDefault = false;

        public const string SubmissionRangeName = "submission";
        public const string HeaderRangeName = "header";
        public const string SublinesRangeName = "sublines";
        public const string SegmentRangeName = "segment";
        public const string UnderwritingYearRangeName = "underwritingYear";
        public const string ProspectiveExposureAmountRangeName = "prospective.exposureAmount";
        public const string PolicyProfileRangeName = "policyProfile";
        public const string HazardProfileRangeName = "hazardProfile";
        public const string StateProfileRangeName = "stateProfile";
        public const string ConstructionTypeProfileRangeName = "constructionTypeProfile";
        public const string OccupancyTypeProfileRangeName = "occupancyTypeProfile";
        public const string ProtectionClassProfileRangeName = "protectionClassProfile";
        public const string TotalInsuredValueProfileRangeName = "totalInsuredValueProfile";
        public const string WorkersCompStateHazardProfileRangeName = "workersCompStateHazardProfile";
        public const string WorkersCompClassCodeProfileRangeName = "workersCompClassCodeProfile";
        public const string WorkersCompStateAttachmentProfileRangeName = "workersCompStateAttachmentProfile";
        public const string MinnesotaRetentionRangeName = "minnesotaRetention";
        public const string SublineAllocationRangeName = "sublineProfile";
        public const string UmbrellaAllocationRangeName = "umbrellaProfile";
        public const string AggregateLossSetRangeName = "aggregateLoss";
        public const string IndividualLossSetRangeName = "individualLoss";
        public const string RateChangeSetRangeName = "rateChange";
        public const string ExposureSetRangeName = "historicalExposures";
        public const string PeriodSetRangeName = "historicalPeriods";
        public const string ProfileBasisRangeName = "basis";
        public const string ThresholdRangeName = "threshold";
        public static double LessThanStandardColumnWidth = 10;
        public static double StandardColumnWidth = 15;
        public static double LossColumnWidth = 20;
        public static double ExtraLargeColumnWidth = 35;
        public static double MarginColumnWidth = 5;
    }
}