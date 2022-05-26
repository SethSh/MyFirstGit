    using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.Models;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for UserPreferencesSelector.xaml
    /// </summary>
    public partial class UserPreferencesDisplayer
    {
        private readonly UserPreferencesViewModel _viewModel;

        private readonly bool _initialIsAdmin;
        private readonly double _initialPaneWidthFactor;
        private readonly PanePositionType _initialPanePosition;

        public UserPreferences UserPreferences { get; set; }
        public DialogResult DialogResult { get; internal set; }


        public UserPreferencesDisplayer(UserPreferencesViewModel viewModel)
        {
            InitializeComponent();
            DialogResult = DialogResult.Cancel;
            _viewModel = viewModel;
            DataContext = _viewModel;

            _initialIsAdmin = _viewModel.IsAdmin;
            _initialPaneWidthFactor = _viewModel.PaneWidthFactor;
            _initialPanePosition = _viewModel.PanePosition;
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_initialIsAdmin != _viewModel.IsAdmin) Globals.Ribbons.SubmissionRibbon.AdminGroup.Visible = _viewModel.IsAdmin;
            if (HasPaneDisplayChanged()) Globals.ThisWorkbook.SetPaneDisplay(_viewModel.PaneWidthFactor, _viewModel.PanePosition);

            UserPreferences = new UserPreferences
            {
                ExposureBasis = _viewModel.ExposureBasis,
                PeriodType = _viewModel.PeriodType,
                ProfileBasisId = _viewModel.ProfileBasisId,
                ShowMyCedents = _viewModel.ShowMyCedents,
                MyCedents = _viewModel.MyCedents,
                ShowMyUnderwriters = _viewModel.ShowMyUnderwriters,
                MyUnderwriters = _viewModel.MyUnderwriters,
                InsertRowCount = _viewModel.InsertRowCount,
                PolicyProfileRowCount = _viewModel.PolicyProfileCount,
                TotalInsuredValueProfileRowCount = _viewModel.TotalInsuredValueProfileCount,
                EvaluationDateType = _viewModel.EvaluationDateType,
                UnderwritingYearType = _viewModel.UnderwritingYearType,
                HistoricalPeriodCount = _viewModel.HistoricalPeriodCount,
                IndividualLossCount = _viewModel.IndividualLossCount,
                RateChangeCount = _viewModel.RateChangeCount,
                WorkersCompClassCodeCount = _viewModel.WorkersCompClassCodeCount,
                WorkersCompStateAttachmentCount = _viewModel.WorkersCompStateAttachmentCount,
                WorkersCompClassCodesIsActive = _viewModel.WorkersCompClassCodesIsActive,
                WorkersCompStateByHazardGroupIsIndependent = _viewModel.WorkersCompStateByHazardGroupIsIndependent,
                SubjectPolicyAlaeTreatment = _viewModel.SubjectPolicyAlaeTreatment,
                SubmissionRibbonSelectionChoice = _viewModel.SubmissionRibbonSelectionChoice,
                AllowSublineFlexibility = _viewModel.AllowSublineFlexibility,
                IsAggregateLossAndAlaeCombined = _viewModel.IsAggregateLossAndAlaeCombined,
                IsAggregatePaidAvailable = _viewModel.IsAggregatePaidAvailable,
                IsIndividualLossAndAlaeCombined = _viewModel.IsIndividualLossAndAlaeCombined,
                IsIndividualPaidAvailable = _viewModel.IsIndividualPaidAvailable,
                IsPolicyLimitAvailable = _viewModel.IsPolicyLimitAvailable,
                IsPolicyAttachmentAvailable = _viewModel.IsPolicyAttachmentAvailable,
                IsAccidentDateAvailable = _viewModel.IsAccidentDateAvailable,
                IsPolicyDateAvailable = _viewModel.IsPolicyDateAvailable,
                IsReportDateAvailable = _viewModel.IsReportDateAvailable,
                IsEventCodeAvailable = _viewModel.IsEventCodeAvailable,
                IsTotalInsuredValueProfileExpanded = _viewModel.IsTotalInsuredValueProfileExpanded,
                UseMeAsUnderwriter = _viewModel.UseMeAsUnderwriter,
                ArePropertyNodesExpanded = _viewModel.ArePropertyNodesExpanded,
                CheckSynchronization = _viewModel.CheckSynchronization,
                IsAdmin = _viewModel.IsAdmin,
                PanePosition = _viewModel.PanePosition,
                PaneWidthFactor = _viewModel.PaneWidthFactor,
                SegmentWorksheetZoom = _viewModel.SegmentWorksheetZoom
            };
            
            DialogResult = DialogResult.OK;
        }

        private bool HasPaneDisplayChanged()
        {
            if (!_initialPaneWidthFactor.IsEqual(_viewModel.PaneWidthFactor)) return true;
            if (_initialPanePosition != _viewModel.PanePosition) return true;
            return false;
        }

        private void PolicyProfileCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.PolicyProfileCountBackground = IsPolicyProfileRowCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        public bool IsPolicyProfileRowCountValid => _viewModel.PolicyProfileCount.IsBetween(_viewModel.PolicyProfileMinimumRowCount,
            _viewModel.IndividualLossMaximumRowCount);

        private void TotalInsuredValueProfileCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.TotalInsuredValueProfileCountBackground = IsTotalInsuredValueProfileRowCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        public bool IsTotalInsuredValueProfileRowCountValid => _viewModel.TotalInsuredValueProfileCount.IsBetween(_viewModel.TotalInsuredValueProfileMinimumRowCount,
            _viewModel.IndividualLossMaximumRowCount);


        public bool IsInsertRowsDefaultCountValid => _viewModel.InsertRowCount.IsBetween(_viewModel.InsertRowCountMinimum,
            _viewModel.InsertRowCountMaximum);

        private void IndividualLossCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.IndividualLossCountBackground = IsIndividualLossCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void CountTextBox_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private bool IsIndividualLossCountValid => _viewModel.IndividualLossCount.IsBetween(_viewModel.IndividualLossMinimumRowCount, _viewModel.IndividualLossMaximumRowCount);
        private bool IsHistoricalPeriodCountValid => _viewModel.HistoricalPeriodCount.IsBetween(_viewModel.HistoricalPeriodMinCount, _viewModel.HistoricalPeriodMaxCount);
        private bool IsRateChangeCountValid => _viewModel.RateChangeCount.IsBetween(_viewModel.RateChangeMinimumRowCount, _viewModel.RateChangeMaximumRowCount);
        private bool IsWorkersCompClassCodeCountValid => _viewModel.WorkersCompClassCodeCount.IsBetween(_viewModel.WorkersCompClassCodeMinimumRowCount, _viewModel.WorkersCompClassCodeMaximumRowCount);
        private bool IsWorkersCompStateAttachmentCountValid => _viewModel.WorkersCompStateAttachmentCount.IsBetween(_viewModel.WorkersCompStateAttachmentMinimumColumnCount, _viewModel.WorkersCompStateAttachmentMaximumColumnCount);
        private void HistoricalPeriodCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.HistoricalPeriodCountBackground = IsHistoricalPeriodCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private bool AreDateSelectionsConsistent()
        {
            switch (_viewModel.PeriodType)
            {
                case 1: return _viewModel.IsAccidentDateAvailable;
                case 2: return _viewModel.IsPolicyDateAvailable;
                case 3: return _viewModel.IsReportDateAvailable;
                default:
                    const string message = "Can't find historical period type";
                    throw new ArgumentOutOfRangeException(message);
            }
        }
        private void IsAccidentDateAvailableCheckBox_OnChecked(object sender, RoutedEventArgs e)
        {
            RenderAccidentDate();
        }

        private void RenderAccidentDate()
        {
            _viewModel.AccidentDateBackground = _viewModel.PeriodType == 1 && !_viewModel.IsAccidentDateAvailable ? Brushes.Red : Brushes.Transparent;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void IsPolicyDateAvailableCheckBox_OnChecked(object sender, RoutedEventArgs e)
        {
            RenderPolicyDate();
        }

        private void RenderPolicyDate()
        {
            _viewModel.PolicyDateBackground = _viewModel.PeriodType == 2 && !_viewModel.IsPolicyDateAvailable ? Brushes.Red : Brushes.Transparent;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }
        private void IsReportDateAvailableCheckBox_OnChecked(object sender, RoutedEventArgs e)
        {
            RenderReportDate();
        }

        private void RenderReportDate()
        {
            _viewModel.ReportDateBackground = _viewModel.PeriodType == 3 && !_viewModel.IsReportDateAvailable ? Brushes.Red : Brushes.Transparent;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private bool GetOkButtonEnabledStatus()
        {
            var isValid = IsInsertRowsDefaultCountValid 
                && IsPolicyProfileRowCountValid
                && IsHistoricalPeriodCountValid 
                && IsIndividualLossCountValid 
                && IsRateChangeCountValid
                && IsWorkersCompClassCodeCountValid
                && AreDateSelectionsConsistent();

            if (isValid)
            {
                _viewModel.OkButtonMessage = string.Empty;
            }
            else
            {
                var sb = new StringBuilder();
                if (!IsInsertRowsDefaultCountValid) sb.AppendLine("Change insert row count");
                if (!IsPolicyProfileRowCountValid) sb.AppendLine($"Change {BexConstants.PolicyProfileName.ToLower()} row count");
                if (!IsHistoricalPeriodCountValid) sb.AppendLine($"Change {BexConstants.PeriodName.ToLower()} row count");
                if (!IsIndividualLossCountValid) sb.AppendLine($"Change {BexConstants.IndividualLossSetName} row count");
                if (!IsRateChangeCountValid) sb.AppendLine($"Change {BexConstants.RateChangeSetName} row count");
                if (!IsWorkersCompClassCodeCountValid) sb.AppendLine($"Change {BexConstants.WorkersCompClassCodeProfileName} row count");
                if (!AreDateSelectionsConsistent())
                {
                    var yearName = HistoricalPeriodTypesFromBex.GetName(_viewModel.PeriodType);
                    var dateFieldName = HistoricalPeriodTypesFromBex.GetDateFieldName(_viewModel.PeriodType);
                    sb.AppendLine($"Change either {BexConstants.PeriodName.ToLower()} <{yearName}> or " +
                                  $"select {BexConstants.IndividualLossSetName.ToLower()} <{dateFieldName}>");
                }

                _viewModel.OkButtonMessage = sb.ToString();
            }

            return isValid;
        }

        
        private void PeriodTypeComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RenderAccidentDate();
            RenderPolicyDate();
            RenderReportDate();
        }

        private void InsertRowCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.InsertRowCountBackground = IsInsertRowsDefaultCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void RateChangeCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.RateChangeCountBackground = IsRateChangeCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void WorkersCompClassCodeCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.WorkersCompClassCodeCountBackground = IsWorkersCompClassCodeCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void WorkersCompStateAttachmentCountTextBox_OnLostFocus(object sender, RoutedEventArgs e)
        {
            _viewModel.WorkersCompStateAttachmentCountBackground = IsWorkersCompStateAttachmentCountValid ? Brushes.Transparent : Brushes.Red;
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }
    }
}
