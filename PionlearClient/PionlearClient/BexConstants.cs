namespace PionlearClient
{
    public static class BexConstants
    {
        public const double WorkbookVersion = 1.3;
        public const string WorkbookPassword = "mlapps";
        public const string SubmissionHeaderRangeName = "submission.header";
        public const string WorkbookVersionRangeName = "workbookVersion";
        public const string RangeName = "Range";

        public const string BexName = "BEX";
        public const string UwpfName = "UWPF";
        public const string KeyDataName = "Key Data";
        public const string RatingAnalysisName = "Rating Analysis";
        public const string RatingAnalysesName = "Rating Analyses";
        public const string InventoryTreeName = "Inventory Tree";

        public const int MaximumNumberOfDataComponents = 6;
        public const string DataValidationTitle = "Data Validation";
        public const string DataQualityControlTitle = "Data Quality Control";
        public const string UploadName = "Upload";
        public const string RebuildName = "Rebuild";
        public const string UpdateName = "Update";
        public const string AttachName = "Attach";
        public const string DecoupleName = "Decouple";
        public const string DecouplingName = "Decoupling";
        public static string UploadPreviewName = $"{UploadName} Preview";

        public const string DefaultPackageName = "A New Package";
        public const string DefaultCedentId = "099999";
        public const string CedentName = "Cedent";

        public const char Dash = '\u2013';
        
        public const string NotFoundErrorCode = "404";

        public const string ApplicationName = "Submission Collector";
        public const string ServerDatabaseName = "Database";
        public const string LedgerName = "Ledger";

        public const string UserPreferencesName = "User Preferences";
        public const string UserPreferencesFileName = "UserPreferences.txt";
        public const string StackTraceLogFileName = "StackTrace.txt";

        public const string EarnedPremiumExposureBasisName = "Premium Earned";
        public const string CalendarAccidentYearPeriodTypeName = "Cal/Acc Year";
        public const string PercentProfileBasisName = "Percent";
        public const int PremiumProfileBasisId = 2;
        public const int StateProfileCountrywideId = 0;
        public const int LossOnlySubjectPolicyAlaeTreatmentId = 2;

        public const string UmbrellaName = "Umbrella";
        public static string UmbrellaTypeName = $"{UmbrellaName} Type";

        public const string PolicyName = "Policy";
        public static string PolicyLimitName = $"{PolicyName} {LimitName}";
        public static string PolicySirName = $"{PolicyName} {SirAttachmentName}";

        public const string TivName = "TIV";
        public static string TivAverageName = $"Average {TivName}";
        public const string ShareName = "Share";
        public const string LimitName = "Limit";
        public const string SirAttachmentName = "SIR/Attachment";
        public const string AttachmentName = "Attachment";
        public const string PercentName = "Percent";
        public const string SublineName = "Subline";
        public const string UnderwritingYearName = "Underwriting Year";
        public const string AnalystName = "Analyst";
        public const string ProspectiveExposureAmountName = "Prospective Exposure Amount";
        public const string DateName = "Date";
        public const string EventCodeName = "Event Code";
        public const string DescriptionName = "Description";
        public static string StartDateName = $"Effective {DateName}";
        public static string EndDateName = $"Expiration {DateName}";
        public static string EvaluationDateName = $"Evaluation {DateName}";
        public const string PackageName = "Submission Package";
        public const string SegmentName = "Submission Segment";
        public const string ProfileName = "Profile";
        public const string PeriodName = "Historical Period";
        public const string PeriodSetName = "Historical Periods";
        public const string ExposureName = "Exposure";
        public static string ExposureSetName = $"{ExposureName} Set";
        public static string AggregateLossSetName = $"{AggregateName} Loss Set";
        public static string AggregateLossSetShortName = $"{AggregateShortName} Loss Set";
        public const string RateChangeName = "Rate Change";
        public static string RateChangeSetName = $"{RateChangeName} Set";
        public const string IndividualLossName = "Individual Loss";
        public static string IndividualLossSetName = $"{IndividualLossName} Set";
        public const string AggregateName = "Aggregate";
        public const string AggregateShortName = "Agg";
        public const string ThresholdName = "Threshold";

        public const string PaidName = "Paid";
        public const string ReportedName = "Reported";
        public const string LossName = "Loss";
        public const string AlaeName = "ALAE";

        public const int AutoLiabilityCommercialSublineCode = 20;
        public const int AutoLiabilityPersonalSublineCode = 21;
        public const int AutoLiabilityPublicEntitySublineCode = 22;

        public const int AutoPhysicalDamageCommercialSublineCode = 29;
        public const int AutoPhysicalDamagePersonalSublineCode = 30;

        public const int GeneralLiabilityHomeownersLiabilitySublineCode = 18;
        public const int GeneralLiabilityHomeownersPropertySublineCode = 24;
        public const int GeneralLiabilityPremOpsSublineCode = 16;
        public const int GeneralLiabilityProductsSublineCode = 17;
        public const int GeneralLiabilityPublicEntitySublineCode = 19;

        public const int CommercialPropertySublineCode = 23;

        public const int ProfessionalLiabilityAttorneysSublineCode = 10;
        public const int ProfessionalLiabilityDirectorsAndOfficersSublineCode = 12;
        public const int ProfessionalLiabilityEmploymentPracticesSublineCode = 11;
        public const int ProfessionalLiabilityHospitalsSublineCode = 7;
        public const int ProfessionalLiabilityMiscellaneousSublineCode = 9;
        public const int ProfessionalLiabilityNursingHomesSublineCode = 8;
        public const int ProfessionalLiabilityPhysiciansSublineCode = 5;
        public const int ProfessionalLiabilityPublicOfficersErrorAndOmissionsSublineCode = 14;
        public const int ProfessionalLiabilitySchoolBoardLegalSublineCode = 13;
        public const int ProfessionalLiabilitySurgeonsSublineCode = 6;
        public const int WorkersCompensationMedicalSublineCode = 38;
        public const int WorkersCompensationIndemnitySublineCode = 4;

        public static string SublineAllocationName = $"{SublineName} {ProfileName}";
        public static string UmbrellaAllocationName = $"Umbrella {ProfileName}";
        public static string PolicyProfileName = $"Policy {ProfileName}";
        public static string StateProfileName = $"State {ProfileName}";
        public static string HazardProfileName = $"Hazard {ProfileName}";
        
        public static string PropertyName = "Property";
        public static string ConstructionTypeName = "Construction Type";
        public static string ConstructionTypeProfileName = $"{ConstructionTypeName} {ProfileName}";
        public static string OccupancyTypeName = "Occupancy Type";
        public static string OccupancyTypeProfileName = $"{OccupancyTypeName} {ProfileName}";
        public static string ProtectionClassName = "Protection Class";
        public static string ProtectionClassProfileName = $"{ProtectionClassName} {ProfileName}";
        public static string TotalInsuredValueProfileName = $"Total Insured Value {ProfileName}";
        public static string TotalInsuredValueAbbreviatedProfileName = $"{TivName} {ProfileName}";

        public static string WorkersCompName = "Workers Comp";
        public static string WorkersCompStateHazardGroupName = $"{WorkersCompName} Hazard Group";
        public static string WorkersCompStateHazardGroupProfileName = $"{WorkersCompStateHazardGroupName} {ProfileName}";
        public static string WorkersCompClassCodeName = $"{WorkersCompName} Class Code";
        public static string WorkersCompClassCodeProfileName = $"{WorkersCompClassCodeName} {ProfileName}";
        public static string WorkersCompStateAttachmentName = $"{WorkersCompName} {AttachmentName}";
        public static string WorkersCompStateAttachmentProfileName = $"{WorkersCompStateAttachmentName} {ProfileName}";
        public static string MinnesotaRetentionName = "Minnesota Retention";

        public static string OccurrenceIdName = "Occurrence ID";
        public static string ClaimIdName = "Claim ID";
        public static string AccidentDateName = $"Accident {DateName}";
        public static string PolicyDateName = $"Policy {DateName}";
        public static string ReportDateName = $"Report {DateName}";
        public static string LossAndAlaeName = $"{LossName} & {AlaeName}";
        public static string ReportedLossName = $"{ReportedName} {LossName}";
        public static string ReportedAlaeName = $"{ReportedName} {AlaeName}";
        public static string ReportedLossAndAlaeName = $"{ReportedName} {LossAndAlaeName}";
        public static string PaidLossName = $"{PaidName} {LossName}";
        public static string PaidAlaeName = $"{PaidName} {AlaeName}";
        public static string PaidLossAndAlaeName = $"{PaidName} {LossAndAlaeName}";

        public static readonly string HalfTab = new string(' ', 4);
        
        public static string NotRecognizedAsANumber = "isn't recognized as a number";
        public static string NotRecognizedAsADate = "isn't recognized as a date";
        public static string NotRecognizedAsAState = "isn't recognized as a state";
    }
}