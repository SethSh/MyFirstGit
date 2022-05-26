using System.ComponentModel;
using Newtonsoft.Json;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace SubmissionCollector.Models.Segment
{
    public interface IIndividualLossSetDescriptor : ILossSetDescriptor
    {
        bool IsPolicyLimitAvailable { get; set; }
        bool IsPolicyAttachmentAvailable { get; set; }
        bool IsAccidentDateAvailable { get; set; }
        bool IsPolicyDateAvailable { get; set; }
        bool IsReportDateAvailable { get; set; }
        bool IsEventCodeAvailable { get; set; }
        void SetToUserPreferences(UserPreferences userPreferences);
    }

    internal class IndividualLossSetDescriptor : ViewModelBase, IIndividualLossSetDescriptor
    {
        [JsonProperty]
        private readonly ISegment _segment;
        
        private bool _isPaidAvailable;
        private bool _isPolicyLimitAvailable;
        private bool _isPolicyAttachmentAvailable;
        private bool _isLossAndAlaeCombined;
        private bool _isAccidentDateAvailable;
        private bool _isPolicyDateAvailable;
        private bool _isReportDateAvailable;
        private bool _isEventCodeAvailable;

        public IndividualLossSetDescriptor(ISegment segment)
        {
            _segment = segment;
        }

        [DisplayName(@"Loss & ALAE Combined")]
        [ReadOnly(false)]
        [PropertyOrder(0)]
        public bool IsLossAndAlaeCombined
        {
            get => _isLossAndAlaeCombined;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyFixing)
                {
                    _isLossAndAlaeCombined = value;
                    NotifyPropertyChanged();
                    return;
                }
                
                _isLossAndAlaeCombined = value;
                NotifyPropertyChanged();
                SetToDirtyDueToCombinedChange(value);
                SetLossesToDirty();
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2. ModifyRangesToReflectChangeToAlaeFormat(_segment);
            }
        }

        [DisplayName("Paid")]
        [Description("Include Paid Column(s)")]
        [ReadOnly(false)]
        [PropertyOrder(1)]
        public bool IsPaidAvailable
        {
            get => _isPaidAvailable;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyFixing)
                {
                    _isPaidAvailable = value;
                    NotifyPropertyChanged();
                    return;
                }

                _isPaidAvailable = value;
                NotifyPropertyChanged();
                SetToDirtyDueToPaidChange(value);
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToPaid(_segment);
            }
        }

        [DisplayName(@"Event Code")]
        [Description("Include Event Code")]
        [ReadOnly(false)]
        [PropertyOrder(2)]
        public bool IsEventCodeAvailable
        {
            get => _isEventCodeAvailable;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyCreating || Segment.IsCurrentlyFixing)
                {
                    _isEventCodeAvailable = value;
                    NotifyPropertyChanged();
                    return;
                }

                _isEventCodeAvailable = value;
                NotifyPropertyChanged();
                SetLossesToDirty(value);
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToEventCode(_segment);
            }
        }

        public void SetToUserPreferences(UserPreferences userPreferences)
        {
            IsLossAndAlaeCombined = userPreferences.IsIndividualLossAndAlaeCombined;
            IsPaidAvailable = userPreferences.IsIndividualPaidAvailable;
            IsPolicyLimitAvailable = userPreferences.IsPolicyLimitAvailable;
            IsPolicyAttachmentAvailable = userPreferences.IsPolicyAttachmentAvailable;
            IsAccidentDateAvailable = userPreferences.IsAccidentDateAvailable;
            IsPolicyDateAvailable = userPreferences.IsPolicyDateAvailable;
            IsReportDateAvailable = userPreferences.IsReportDateAvailable;
            IsEventCodeAvailable = userPreferences.IsEventCodeAvailable;
        }

        [DisplayName(@"Policy Limit")]
        [Description("Include Policy Limit Column")]
        [ReadOnly(false)]
        [PropertyOrder(3)]
        public bool IsPolicyLimitAvailable
        {
            get => _isPolicyLimitAvailable;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyFixing)
                {
                    _isPolicyLimitAvailable = value;
                    NotifyPropertyChanged();
                    return;
                }

                _isPolicyLimitAvailable = value;
                NotifyPropertyChanged();
                SetLossesToDirty(value);
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToLimit(_segment);
            }
        }

        [DisplayName(@"Policy SIR/Attachment")]
        [Description("Include Policy SIR/Attachment Column")]
        [ReadOnly(false)]
        [PropertyOrder(4)]
        public bool IsPolicyAttachmentAvailable
        {
            get => _isPolicyAttachmentAvailable;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyFixing)
                {
                    _isPolicyAttachmentAvailable = value;
                    NotifyPropertyChanged();
                    return;
                }

                _isPolicyAttachmentAvailable = value;
                NotifyPropertyChanged();
                SetLossesToDirty(value);
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToAttachment(_segment);
            }
        }

        [DisplayName(@"Accident Date")]
        [Description("Include Accident Date")]
        [ReadOnly(false)]
        [PropertyOrder(5)]
        public bool IsAccidentDateAvailable
        {
            get => _isAccidentDateAvailable;
            set
            {

                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyCreating || Segment.IsCurrentlyFixing)
                {
                    _isAccidentDateAvailable = value;
                    NotifyPropertyChanged();
                    return;
                }
                
                if (!value && !IsPolicyDateAvailable && !IsReportDateAvailable)
                {
                    MessageHelper.Show("Can't uncheck the last date field", MessageType.Stop);
                    return;
                }

                _isAccidentDateAvailable = value;
                NotifyPropertyChanged();
                SetLossesToDirty(value);
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToAccidentDate(_segment);
            }
        }

        [DisplayName(@"Policy Date")]
        [Description("Include Policy Date Column")]
        [ReadOnly(false)]
        [PropertyOrder(6)]
        public bool IsPolicyDateAvailable
        {
            get => _isPolicyDateAvailable;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyCreating || Segment.IsCurrentlyFixing)
                {
                    _isPolicyDateAvailable = value;
                    NotifyPropertyChanged();
                    return;
                }

                if (!value && !IsAccidentDateAvailable && !IsReportDateAvailable)
                {
                    MessageHelper.Show("Can't uncheck the last date field", MessageType.Stop);
                    return;
                }

                _isPolicyDateAvailable = value;
                NotifyPropertyChanged();
                SetLossesToDirty(value);
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToPolicyDate(_segment);
            }
        }

        [DisplayName(@"Report Date")]
        [Description("Include Report Date")]
        [ReadOnly(false)]
        [PropertyOrder(7)]
        public bool IsReportDateAvailable
        {
            get => _isReportDateAvailable;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || Segment.IsCurrentlyDuplicating || Segment.IsCurrentlyCreating || Segment.IsCurrentlyFixing)
                {
                    _isReportDateAvailable = value;
                    NotifyPropertyChanged();
                    return;
                }

                if (!value && !IsAccidentDateAvailable && !IsPolicyDateAvailable)
                {
                    MessageHelper.Show("Can't uncheck the last date field", MessageType.Stop);
                    return;
                }

                _isReportDateAvailable = value;
                NotifyPropertyChanged();
                SetLossesToDirty(value);
                if (_segment.WorksheetManager != null) IndividualLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToReportDate(_segment);
            }
        }
        
        public override string ToString()
        {
            return string.Empty;
        }

        private void SetToDirtyDueToPaidChange(bool isSelected)
        {
            _segment.IsDirty = true;
            SetLossesToDirty(isSelected);
        }
        
        private void SetToDirtyDueToCombinedChange(bool isSelected)
        {
            SetLossSetsToDirty();
            SetLossesToDirty(isSelected);
        }

        private void SetLossSetsToDirty()
        {
            foreach (var lossSet in _segment.IndividualLossSets)
            {
                lossSet.IsDirty = true;
            }
        }

        private void SetLossesToDirty(bool isSelected)
        {
            //selecting doesn't change losses, but de-selecting does (might)
            if (!isSelected)
            {
                SetLossesToDirty();
            }
        }

        private void SetLossesToDirty()
        {
            foreach (var lossSet in _segment.IndividualLossSets)
            {
                lossSet.Ledger.SetToDirty();
            }
        }
    }
}