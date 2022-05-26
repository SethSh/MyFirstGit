using System.ComponentModel;
using Newtonsoft.Json;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.ViewModel;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace SubmissionCollector.Models.Segment
{
    public interface IAggregateLossSetDescriptor : ILossSetDescriptor
    {
        void SetToUserPreferences(UserPreferences userPreferences);
    }

    public class AggregateLossSetDescriptor : ViewModelBase, IAggregateLossSetDescriptor
    {
        [JsonProperty]
        private readonly ISegment _segment;
        
        private bool _isLossAndAlaeCombined;
        private bool _isPaidAvailable;


        public AggregateLossSetDescriptor(ISegment segment)
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
                if (_segment.WorksheetManager != null) AggregateLossSetExcelMatrixHelper2.ModifyRangesToReflectChangeToAlaeFormat(_segment);
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
                if (_segment.WorksheetManager != null) AggregateLossSetExcelMatrixHelper2.ModifyRangesToReflectPaids(_segment);
            }
        }

        public void SetToUserPreferences(UserPreferences userPreferences)
        {
            IsLossAndAlaeCombined = userPreferences.IsAggregateLossAndAlaeCombined;
            IsPaidAvailable = userPreferences.IsAggregatePaidAvailable;
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
            foreach (var lossSet in _segment.AggregateLossSets)
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
            foreach (var lossSet in _segment.AggregateLossSets)
            {
                lossSet.Ledger.SetToDirty();
            }
        }
    }
}