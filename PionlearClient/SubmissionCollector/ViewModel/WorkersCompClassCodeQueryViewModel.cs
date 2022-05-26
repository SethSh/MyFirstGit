using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PionlearClient.BexReferenceData;

namespace SubmissionCollector.ViewModel
{

    internal class WorkersCompClassCodeSearchViewModel : BaseWorkersCompClassCodeSearchViewModel
    {
        public WorkersCompClassCodeSearchViewModel()
        {
            var stateClassCodeSets = WorkersCompClassCodesAndHazardsFromBex.StateClassCodes.ToList();
            var hazardGroups = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.ToDictionary(key => key.Id);

            foreach (var codeSet in stateClassCodeSets.OrderBy(codeSet => codeSet.State.Abbreviation))
            {
                var state = codeSet.State;
                foreach (var classCode in codeSet.ClassCodes.OrderBy(classCode => classCode.StateClassCode))
                {
                    ClassCodeViewItems.Add(new WorkersCompClassCodeQueryViewItem
                    {
                        StateAbbreviation = state.Abbreviation,
                        StateClassCode = classCode.StateClassCode,
                        HazardGroupName = classCode.HazardGroupId.HasValue ? hazardGroups[classCode.HazardGroupId.Value].Name : string.Empty,
                        StateDescription = classCode.StateDescription
                    });
                }
            }
        }
    }

    internal abstract class BaseWorkersCompClassCodeSearchViewModel : ViewModelBase
    {
        private string _queryCriteria;
        private IList<WorkersCompClassCodeQueryViewItem> _filteredClassCodeViewItems;
        private bool _isSearching;
        private string _statusMessage;
        private int _itemCount;
        protected BaseWorkersCompClassCodeSearchViewModel()
        {
            ClassCodeViewItems = new List<WorkersCompClassCodeQueryViewItem>();
            SearchCriteria = string.Empty;
        }
        
        public IList<WorkersCompClassCodeQueryViewItem> ClassCodeViewItems { get; set; }
        public IList<WorkersCompClassCodeQueryViewItem> FilteredClassCodeViewItems
        {
            get  => _filteredClassCodeViewItems;
            set
            {
                _filteredClassCodeViewItems = value;
                NotifyPropertyChanged();
            }
        }

        public string SearchCriteria
        {
            get => _queryCriteria;
            set
            {
                _queryCriteria = value;
                NotifyPropertyChanged();
                SetFilter();
            }
        }
        
        public bool IsSearching
        {
            get => _isSearching;
            set
            {
                _isSearching = value;
                NotifyPropertyChanged();
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                _statusMessage = value;
                NotifyPropertyChanged();
            }
        }
        
        public int ItemCount
        {
            get=>_itemCount;
            set
            {
                _itemCount = value;
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("ItemCountLabel");
            }
        }

        public string ItemCountLabel => ItemCount > 0 ? $"{ItemCount:N0} matches" : string.Empty;

        
        private IList<WorkersCompClassCodeQueryViewItem> GetMatchingItems(string queryCriteria)
        {
            var matches = int.TryParse(queryCriteria, out var criteriaAsInteger) 
                ? ClassCodeViewItems.Where(item => item.StateClassCode == criteriaAsInteger) 
                : ClassCodeViewItems.Where(item => item.StateDescription.IndexOf(queryCriteria, StringComparison.OrdinalIgnoreCase) > -1);

            return matches.ToList();
        }

        private void SetFilter()
        {
            IsSearching = true;

            FilteredClassCodeViewItems = new List<WorkersCompClassCodeQueryViewItem>();

            if (!ClassCodeViewItems.Any() || string.IsNullOrEmpty(SearchCriteria))
            {
                IsSearching = false;
                StatusMessage = string.Empty;
                ItemCount = 0;
                return;
            }
            
            StatusMessage = "Search in progress ...";

            var cancellationTokenSource = new CancellationTokenSource();

            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            var task = new Task<IList<WorkersCompClassCodeQueryViewItem>>(() => GetMatchingItems(SearchCriteria));
            task.ContinueWith(task1 =>
            {
                if (task1.IsFaulted)
                {
                    if (task1.Exception?.InnerException != null)
                    {
                        StatusMessage = task1.Exception.InnerException.Message;
                    }
                }

                FilteredClassCodeViewItems = task.Result;
                IsSearching = false;
                StatusMessage = string.Empty;
                ItemCount = FilteredClassCodeViewItems.Count;

            }, cancellationTokenSource.Token, TaskContinuationOptions.None, scheduler);

            task.Start();
        }
    }

    internal class WorkersCompClassCodeQueryViewItem
    {
        public string StateAbbreviation { get; set; }
        public long StateClassCode { get; set; }
        public string StateClassCodeAsString => StateClassCode.ToString("0000");
        public string HazardGroupName { get; set; }
        public string StateDescription { get; set; }
    }
}