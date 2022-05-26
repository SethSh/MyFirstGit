using System.ComponentModel;

namespace SubmissionCollector.Models
{
    public abstract class BaseInventoryItem : BaseModelProxy, IInventoryItem
    {
        private bool _isSelected;
        private bool _isExpanded;
        private int _displayOrder;
        
        [Category(@"Development")]
        [DisplayName(@"Order Id")]
        [Browsable(false)]
        [ReadOnly(true)]
        public int DisplayOrder
        {
            get => _displayOrder;
            set
            {
                _displayOrder = value;
                NotifyPropertyChanged();
            }
        }

        [Category(@"Development")]
        [Browsable(false)]
        [ReadOnly(true)]
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                NotifyPropertyChanged();
            }
        }

        [Category(@"Development")]
        [Browsable(false)]
        [ReadOnly(true)]
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                NotifyPropertyChanged();
            }
        }
    }
}