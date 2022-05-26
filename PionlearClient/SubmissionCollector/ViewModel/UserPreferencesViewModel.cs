using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Windows.Data;
using System.Windows.Media;
using MunichRe.Bex.ApiClient.ClientApi;
using PionlearClient.BexReferenceData;
using PionlearClient.Enums;
using PionlearClient.Extensions;
using PionlearClient.KeyDataFolder;

namespace SubmissionCollector.ViewModel
{
    public class UserPreferencesViewModel : BaseUserPreferencesViewModel
    {
        public UserPreferencesViewModel(UserPreferences up)
        {
            ExposureBasis = up.ExposureBasis;
            PeriodType = up.PeriodType;
            ProfileBasisId = up.ProfileBasisId;
            ShowMyCedents = up.ShowMyCedents;
            MyCedents = up.MyCedents;
            ShowMyUnderwriters = up.ShowMyUnderwriters;
            MyUnderwriters = up.MyUnderwriters;
            InsertRowCount = up.InsertRowCount;
            PolicyProfileCount = up.PolicyProfileRowCount;
            PolicyProfileCountBackground = Brushes.Transparent;
            TotalInsuredValueProfileCount = up.TotalInsuredValueProfileRowCount;
            TotalInsuredValueProfileCountBackground = Brushes.Transparent;
            EvaluationDateType = up.EvaluationDateType;
            UnderwritingYearType = up.UnderwritingYearType;
            HistoricalPeriodCount = up.HistoricalPeriodCount;
            IndividualLossCount = up.IndividualLossCount;
            RateChangeCount = up.RateChangeCount;
            WorkersCompClassCodeCount = up.WorkersCompClassCodeCount;
            WorkersCompStateAttachmentCount = up.WorkersCompStateAttachmentCount;
            WorkersCompClassCodesIsActive = up.WorkersCompClassCodesIsActive;
            WorkersCompStateByHazardGroupIsIndependent = up.WorkersCompStateByHazardGroupIsIndependent;
            SubjectPolicyAlaeTreatment = up.SubjectPolicyAlaeTreatment;
            SubmissionRibbonSelectionChoice = up.SubmissionRibbonSelectionChoice;
            AllowSublineFlexibility = up.AllowSublineFlexibility;
            IsAggregateLossAndAlaeCombined = up.IsAggregateLossAndAlaeCombined;
            IsAggregatePaidAvailable= up.IsAggregatePaidAvailable;
            IsIndividualLossAndAlaeCombined = up.IsIndividualLossAndAlaeCombined;
            IsIndividualPaidAvailable = up.IsIndividualPaidAvailable;
            IsPolicyLimitAvailable = up.IsPolicyLimitAvailable;
            IsPolicyAttachmentAvailable = up.IsPolicyAttachmentAvailable;
            IsAccidentDateAvailable = up.IsAccidentDateAvailable;
            IsPolicyDateAvailable = up.IsPolicyDateAvailable;
            IsReportDateAvailable = up.IsReportDateAvailable;
            IsEventCodeAvailable = up.IsEventCodeAvailable;
            IsTotalInsuredValueProfileExpanded = up.IsTotalInsuredValueProfileExpanded;
            UseMeAsUnderwriter = up.UseMeAsUnderwriter;
            ArePropertyNodesExpanded = up.ArePropertyNodesExpanded;
            CheckSynchronization = up.CheckSynchronization;
            OkButtonEnabled = true;
            OkButtonMessage = string.Empty;
            IsAdmin = up.IsAdmin;
            PaneWidthFactor = up.PaneWidthFactor;
            PanePosition = up.PanePosition;
            SegmentWorksheetZoom = up.SegmentWorksheetZoom;

            MyCedents.ForEach(x => x.Id = x.ClippedId);

            ExposureBases = new CollectionView(ExposureBasisFromBex.ReferenceData);
            PeriodTypes = new CollectionView(HistoricalPeriodTypesFromBex.ReferenceData);
            ProfileBases = new CollectionView(ProfileBasisFromBex.ReferenceData);
        }
    }

    public class BaseUserPreferencesViewModel : ViewModelBase
    {
        private int _exposureBasis;
        private int _periodType;
        private int _profileBasisId;
        private CollectionView _profileBases;
        private CollectionView _exposureBases;
        private CollectionView _periodTypes;
        private bool _showMyCedents;
        private IList<BusinessPartner> _myCedents;
        private bool _showMyUnderwriters;
        private IList<Underwriter> _myUnderwriters;
        private int _policyProfileCount;
        private SolidColorBrush _policyProfileCountBackground;
        private SolidColorBrush _accidentDateBackground;
        private SolidColorBrush _policyDateBackground;
        private SolidColorBrush _reportDateBackground;
        private bool _okButtonEnabled;
        private EvaluationDateType _evaluationDateType;
        private SolidColorBrush _historicalPeriodCountBackground;
        private bool _submissionRibbonSelectionChoice;
        private int _historicalPeriodCount;
        private int _individualLossCount;
        private int _rateChangeCount;
        private int _workersCompClassCodeCount;
        private int _workersCompStateAttachmentCount;
        private bool _workersCompClassCodesIsActive;
        private SolidColorBrush _individualLossCountBackground;
        private SolidColorBrush _rateChangeCountBackground;
        private SolidColorBrush _workersCompClassCodeCountBackground;
        private SolidColorBrush _workersCompStateAttachmentCountBackground;
        private bool _allowSublineFlexibility;
        private bool _isAggregateLossAndAlaeCombined;
        private bool _isAggregatePaidAvailable;
        private bool _isIndividualLossAndAlaeCombined;
        private bool _isIndividualPaidAvailable;
        private bool _isPolicyLimitAvailable;
        private bool _isPolicyAttachmentAvailable;
        private UnderwritingYearType _underwritingYearType;
        private bool _useMeAsUnderwriter;
        private bool _isAccidentDateAvailable;
        private bool _isPolicyDateAvailable;
        private bool _isReportDateAvailable;
        private string _okButtonMessage;
        private bool _arePropertyNodesExpanded;
        private bool _checkSynchronization;
        private bool _isAdmin;
        private SolidColorBrush _insertRowCountBackground;
        private int _insertRowCount;
        private bool _isEventCodeAvailable;
        private PanePositionType _panePosition;
        private double _paneWidthFactor;
        private short _segmentWorksheetZoom;
        private int _totalInsuredValueProfileCount;
        private SolidColorBrush _totalInsuredValueProfileCountBackground;
        private bool _isTotalInsuredValueProfileExpanded;
        private bool _workersCompStateByHazardGroupIsIndependent;

        public int TotalInsuredValueProfileMaximumRowCount => 250;
        public int TotalInsuredValueProfileMinimumRowCount => 10;
        public string TotalInsuredValueProfileMaximumRowCountMessage => $"Maximum value is {TotalInsuredValueProfileMaximumRowCount:N0}";
        public string TotalInsuredValueProfileMinimumRowCountMessage => $"Minimum value is {TotalInsuredValueProfileMinimumRowCount:N0}";

        public int IndividualLossMaximumRowCount => 2500;
        public int IndividualLossMinimumRowCount => 5;
        public string IndividualLossMaximumRowCountMessage => $"Maximum value is {IndividualLossMaximumRowCount}";
        public string IndividualLossMinimumRowCountMessage => $"Minimum value is {IndividualLossMinimumRowCount}";

        public int RateChangeMaximumRowCount => 2500;
        public int RateChangeMinimumRowCount => 5;
        public string RateChangeMaximumRowCountMessage => $"Maximum value is {RateChangeMaximumRowCount}";
        public string RateChangeMinimumRowCountMessage => $"Minimum value is {RateChangeMinimumRowCount}";



        public int WorkersCompClassCodeMaximumRowCount => 2500;
        public int WorkersCompClassCodeMinimumRowCount => 5;
        public string WorkersCompClassCodeMaximumRowCountMessage => $"Maximum value is {WorkersCompClassCodeMaximumRowCount}";
        public string WorkersCompClassCodeMinimumRowCountMessage => $"Minimum value is {WorkersCompClassCodeMinimumRowCount}";

        public int WorkersCompStateAttachmentMaximumColumnCount => 10;
        public int WorkersCompStateAttachmentMinimumColumnCount => 1;
        public string WorkersCompStateAttachmentMaximumColumnCountMessage => $"Maximum value is {WorkersCompStateAttachmentMaximumColumnCount}";
        public string WorkersCompStateAttachmentMinimumColumnCountMessage => $"Minimum value is {WorkersCompStateAttachmentMinimumColumnCount}";


        public string Title => $"{UserPrincipal.Current.GivenName} {UserPrincipal.Current.Surname} User Preferences";
        public EvaluationDateType EvaluationDateType
        {
            get => _evaluationDateType;
            set
            {
                _evaluationDateType = value;
                NotifyPropertyChanged();
            }
        }

        public UnderwritingYearType UnderwritingYearType
        {
            get => _underwritingYearType;
            set
            {
                _underwritingYearType = value;
                NotifyPropertyChanged();
            }
        }

        public int ProfileBasisId
        {
            get => _profileBasisId;
            set
            {
                _profileBasisId = value;
                NotifyPropertyChanged();
            }
        }
        public int SubjectPolicyAlaeTreatment { get; set; }
        public IEnumerable<double> PaneWidthFactors => new List<double> {0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5 };
        public IEnumerable<short> SegmentWorksheetZooms => new List<short> { 75, 80, 85, 90, 95, 100, 105, 110, 115, 120, 125, 130, 135, 140, 145, 150 };
        public IEnumerable<PanePositionType> PanePositionTypes => Enum.GetValues(typeof(PanePositionType)).Cast<PanePositionType>();
        public IEnumerable<EvaluationDateType> EvaluationDateTypes => Enum.GetValues(typeof(EvaluationDateType)).Cast<EvaluationDateType>();
        public IEnumerable<UnderwritingYearType> UnderwritingYearTypes => Enum.GetValues(typeof(UnderwritingYearType)).Cast<UnderwritingYearType>();

        public IEnumerable<SubjectPolicyAlaeTreatmentViewModel> SubjectPolicyAlaeTreatments => SubjectPolicyAlaeTreatmentsFromBex.ReferenceData;
        public DateTime TodaysDate => DateTime.Now;
        public DateTime LastDayOfLastQuarter => GetLastDayOfLastQuarter();
        public DateTime LastDayOfLastHalfYear => GetLastDayOfLastHalfYear();
        public DateTime LastDayOfLastYear => GetLastDayOfLastYear();

        public bool SubmissionRibbonSelectionChoice
        {
            get => _submissionRibbonSelectionChoice;
            set
            {
                _submissionRibbonSelectionChoice = value;
                NotifyPropertyChanged();
            }
        }

        public string OkButtonMessage
        {
            get => _okButtonMessage;
            set
            {
                _okButtonMessage = value;
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("OkButtonMessageVisibility");
            }
        }

        public bool OkButtonMessageVisibility => !string.IsNullOrEmpty(OkButtonMessage);
        
        public int ExposureBasis
        {
            get => _exposureBasis;
            set
            {
                _exposureBasis = value;
                NotifyPropertyChanged();
            }
        }

        public int InsertRowCount
        {
            get => _insertRowCount;
            set
            {
                _insertRowCount = value;
                NotifyPropertyChanged();
            }
        }
        
        public int PolicyProfileCount
        {
            get => _policyProfileCount;
            set
            {
                _policyProfileCount = value;
                NotifyPropertyChanged();
            }
        }

        public int TotalInsuredValueProfileCount
        {
            get => _totalInsuredValueProfileCount;
            set
            {
                _totalInsuredValueProfileCount = value;
                NotifyPropertyChanged();
            }
        }
        public int InsertRowCountMaximum => 10000;
        public int InsertRowCountMinimum => 1;
        public string InsertRowCountMaximumMessage => $"Maximum value is {InsertRowCountMaximum:N0}";
        public string InsertRowCountMinimumMessage => $"Minimum value is {InsertRowCountMinimum:N0}";

        public int PolicyProfileMaximumRowCount => 250;
        public int PolicyProfileMinimumRowCount => 10;
        public string PolicyProfileMaximumRowCountMessage => $"Maximum value is {PolicyProfileMaximumRowCount:N0}";
        public string PolicyProfileMinimumRowCountMessage => $"Minimum value is {PolicyProfileMinimumRowCount:N0}";

        public int HistoricalPeriodCount
        {
            get => _historicalPeriodCount;
            set
            {
                _historicalPeriodCount = value;
                NotifyPropertyChanged();
            }
        }

        public int HistoricalPeriodMaxCount => 250;
        public int HistoricalPeriodMinCount => 5;
        public string HistoricalPeriodMaxCountMessage => $"Maximum value is {HistoricalPeriodMaxCount}";
        public string HistoricalPeriodMinCountMessage => $"Minimum value is {HistoricalPeriodMinCount}";

        public int IndividualLossCount
        {
            get => _individualLossCount;
            set
            {
                _individualLossCount = value;
                NotifyPropertyChanged();
            }
        }

        public int RateChangeCount
        {
            get => _rateChangeCount;
            set
            {
                _rateChangeCount = value;
                NotifyPropertyChanged();
            }
        }

        public int WorkersCompClassCodeCount
        {
            get => _workersCompClassCodeCount;
            set
            {
                _workersCompClassCodeCount = value;
                NotifyPropertyChanged();
            }
        }

        public int WorkersCompStateAttachmentCount
        {
            get => _workersCompStateAttachmentCount;
            set
            {
                _workersCompStateAttachmentCount = value;
                NotifyPropertyChanged();
            }
        }

        public bool WorkersCompClassCodesIsActive
        {
            get => _workersCompClassCodesIsActive;
            set
            {
                _workersCompClassCodesIsActive = value;
                NotifyPropertyChanged();
            }
        }

        public bool WorkersCompStateByHazardGroupIsIndependent
        {
            get => _workersCompStateByHazardGroupIsIndependent;
            set
            {
                _workersCompStateByHazardGroupIsIndependent = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush PolicyProfileCountBackground
        {
            get => _policyProfileCountBackground;
            set
            {
                _policyProfileCountBackground = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush TotalInsuredValueProfileCountBackground
        {
            get => _totalInsuredValueProfileCountBackground;
            set
            {
                _totalInsuredValueProfileCountBackground = value;
                NotifyPropertyChanged();
            }
        }
        public SolidColorBrush AccidentDateBackground
        {
            get => _accidentDateBackground;
            set
            {
                _accidentDateBackground = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush PolicyDateBackground
        {
            get => _policyDateBackground;
            set
            {
                _policyDateBackground = value;
                NotifyPropertyChanged();
            }
        }
        public SolidColorBrush ReportDateBackground
        {
            get => _reportDateBackground;
            set
            {
                _reportDateBackground = value;
                NotifyPropertyChanged();
            }
        }
        public SolidColorBrush HistoricalPeriodCountBackground
        {
            get => _historicalPeriodCountBackground;
            set
            {
                _historicalPeriodCountBackground = value; 
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush IndividualLossCountBackground
        {
            get => _individualLossCountBackground;
            set
            {
                _individualLossCountBackground = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush RateChangeCountBackground
        {
            get => _rateChangeCountBackground;
            set
            {
                _rateChangeCountBackground = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush WorkersCompClassCodeCountBackground
        {
            get => _workersCompClassCodeCountBackground;
            set
            {
                _workersCompClassCodeCountBackground = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush WorkersCompStateAttachmentCountBackground
        {
            get => _workersCompStateAttachmentCountBackground;
            set
            {
                _workersCompStateAttachmentCountBackground = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush InsertRowCountBackground
        {
            get => _insertRowCountBackground;
            set
            {
                _insertRowCountBackground = value;
                NotifyPropertyChanged();
            }
        }

        public bool OkButtonEnabled
        {
            get => _okButtonEnabled;
            set
            {
                _okButtonEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public IList<BusinessPartner> MyCedents
        {
            get => _myCedents;
            set
            {
                _myCedents = value;
                NotifyPropertyChanged();
            }
        }

        public IList<Underwriter> MyUnderwriters
        {
            get => _myUnderwriters;
            set
            {
                _myUnderwriters = value;
                NotifyPropertyChanged();
            }
        }

        public int PeriodType
        {
            get => _periodType;
            set
            {
                _periodType = value;
                NotifyPropertyChanged();
            }
        }

        public CollectionView ExposureBases
        {
            get => _exposureBases;
            set
            {
                _exposureBases = value;
                NotifyPropertyChanged();
            }
        }

        public CollectionView ProfileBases
        {
            get => _profileBases;
            set
            {
                _profileBases = value;
                NotifyPropertyChanged();
            }
        }

        public CollectionView PeriodTypes
        {
            get => _periodTypes;
            set
            {
                _periodTypes = value;
                NotifyPropertyChanged();
            }
        }

        public bool ShowMyUnderwriters
        {
            get => _showMyUnderwriters;
            set
            {
                _showMyUnderwriters = value;
                NotifyPropertyChanged();
            }
        }

        public bool ShowMyCedents
        {
            get => _showMyCedents;
            set
            {
                _showMyCedents = value;
                NotifyPropertyChanged();
            }
        }

        public bool AllowSublineFlexibility
        {
            get => _allowSublineFlexibility;
            set
            {
                _allowSublineFlexibility = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsAggregateLossAndAlaeCombined
        {
            get => _isAggregateLossAndAlaeCombined;
            set
            {
                _isAggregateLossAndAlaeCombined = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsAggregatePaidAvailable
        {
            get => _isAggregatePaidAvailable;
            set
            {
                _isAggregatePaidAvailable = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsIndividualLossAndAlaeCombined
        {
            get => _isIndividualLossAndAlaeCombined;
            set
            {
                _isIndividualLossAndAlaeCombined = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsIndividualPaidAvailable
        {
            get => _isIndividualPaidAvailable;
            set
            {
                _isIndividualPaidAvailable = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsPolicyLimitAvailable
        {
            get => _isPolicyLimitAvailable;
            set
            {
                _isPolicyLimitAvailable = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsPolicyAttachmentAvailable
        {
            get => _isPolicyAttachmentAvailable;
            set
            {
                _isPolicyAttachmentAvailable = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsAccidentDateAvailable
        {
            get => _isAccidentDateAvailable;
            set
            {
                _isAccidentDateAvailable = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsPolicyDateAvailable
        {
            get => _isPolicyDateAvailable;
            set
            {
                _isPolicyDateAvailable = value;
                NotifyPropertyChanged();
            }
        }
        public bool IsReportDateAvailable
        {
            get => _isReportDateAvailable;
            set
            {
                _isReportDateAvailable = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsEventCodeAvailable
        {
            get => _isEventCodeAvailable;
            set
            {
                _isEventCodeAvailable = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsTotalInsuredValueProfileExpanded
        {
            get => _isTotalInsuredValueProfileExpanded;
            set
            {
                _isTotalInsuredValueProfileExpanded = value;
                NotifyPropertyChanged();
            }
        }
        public bool UseMeAsUnderwriter
        {
            get => _useMeAsUnderwriter;
            set
            {
                _useMeAsUnderwriter = value; 
                NotifyPropertyChanged();
            }
        }

        public bool ArePropertyNodesExpanded
        {
            get => _arePropertyNodesExpanded;
            set
            {
                _arePropertyNodesExpanded = value;
                NotifyPropertyChanged();
            }
        }

        public bool CheckSynchronization
        {
            get => _checkSynchronization;
            set
            {
                _checkSynchronization = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsAdmin
        {
            get => _isAdmin;
            set
            {
                _isAdmin = value;
                NotifyPropertyChanged();
            }
        }

        internal static DateTime GetLastDayOfLastQuarter()
        {
            var date = DateTime.Now;
            
            if (date.Month < 4)
            {
                return new DateTime(date.Year - 1, 12, 31);
            }

            if (date.Month >= 4 && date.Month <= 6)
            {
                return new DateTime(date.Year, 3, 31);
            }

            if (date.Month >= 7 && date.Month <= 9)
            {
                return new DateTime(date.Year, 6, 30);
            }

            return new DateTime(date.Year, 9, 30);
        }

        public double PaneWidthFactor
        {
            get => _paneWidthFactor;
            set
            {
                _paneWidthFactor = value;
                NotifyPropertyChanged();
            }
        }

        public short SegmentWorksheetZoom
        {
            get => _segmentWorksheetZoom;
            set
            {
                _segmentWorksheetZoom = value;
                NotifyPropertyChanged();
            }
        }

        public PanePositionType PanePosition
        {
            get => _panePosition;
            set
            {
                _panePosition = value;
                NotifyPropertyChanged();
            }
        }

        internal static DateTime GetLastDayOfLastHalfYear()
        {
            var date = DateTime.Now;

            if (date.Month < 7)
            {
                return new DateTime(date.Year - 1, 12, 31);
            }

            return new DateTime(date.Year, 6, 30);
        }

        internal static DateTime GetLastDayOfLastYear()
        {
            return new DateTime(DateTime.Now.Year - 1, 12, 31);
        }
    }

    [TypeConverter(typeof(EnumCamelCaseConverter))]
    public enum EvaluationDateType 
    {
        None,
        EndOfLastQuarter,
        EndOfLastHalfYear,
        EndOfLastYear,
    }

    [TypeConverter(typeof(EnumCamelCaseConverter))]
    public enum UnderwritingYearType
    {
        None,
        CurrentYear,
        NextYear
    }

}
