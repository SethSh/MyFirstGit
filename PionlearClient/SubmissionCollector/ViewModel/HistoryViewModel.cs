using System;
using System.Collections.Generic;
using System.Linq;
using SubmissionCollector.Models.Package;

namespace SubmissionCollector.ViewModel
{
    public class HistoryViewModel : BaseHistoryViewModel
    {
        public HistoryViewModel(IPackage package)
        {
            Items = package.BexCommunications.OrderByDescending(comm => comm.Timestamp).Select(
                comm => (IHistoryItemViewModel)new HistoryItemViewModel
                {
                    UserName = comm.UserName,
                    Timestamp = comm.Timestamp,
                    Activity = comm.Activity
                }).ToList();
        }
    }

    public abstract class BaseHistoryViewModel: ViewModelBase
    {
        private ICollection<IHistoryItemViewModel> _items;

        public ICollection<IHistoryItemViewModel> Items
        {
            get => _items;
            set
            {
                _items = value;
                NotifyPropertyChanged();
            }
        }

        protected BaseHistoryViewModel()
        {
            Items = new List<IHistoryItemViewModel>();
        }
    }

    public class HistoryItemViewModel : ViewModelBase, IHistoryItemViewModel
    {
        private string _userName;
        private DateTime _timestamp;
        private string _activity;

        public string UserName
        {
            get => _userName;
            set
            {
                _userName = value; 
                NotifyPropertyChanged();
            }
        }

        public DateTime Timestamp
        {
            get => _timestamp;
            set
            {
                _timestamp = value;
                NotifyPropertyChanged();
            }
        }

        public string Activity
        {
            get => _activity;
            set
            {
                _activity = value;
                NotifyPropertyChanged();
            }
        }
    }

    public interface IHistoryItemViewModel
    {
        string UserName { get; set; }
        DateTime Timestamp { get; set; }
        string Activity { get; set; }
    }
}
