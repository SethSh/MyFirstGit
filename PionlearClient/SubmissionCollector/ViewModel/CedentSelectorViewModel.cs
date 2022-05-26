using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using PionlearClient.KeyDataFolder;
using SubmissionCollector.Enums;
using SubmissionCollector.View;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ViewModel
{
    public class CedentSelectorViewModel : BaseCedentSelectorViewModel
    {
        public CedentSelectorViewModel()
        {
            var up = UserPreferences.ReadFromFile();
            ShowMyCedents = up.ShowMyCedents;
        }
    }

    public abstract class BaseCedentSelectorViewModel : ViewModelBase
    {
        public readonly GridLength CriteriaRowLengthWhenVisible = new GridLength(40);
        public readonly GridLength StatusRowLengthWhenVisible = new GridLength(40);
        public readonly GridLength NoLength = new GridLength(0);
        public readonly int BusinessPartnerCodeLength = 10;
        public readonly string CedentMatchesString = "cedent match(es)";

        private string _statusMessage;
        private bool _showMyCedents;
        private GridLength _criteriaRowLength;
        private GridLength _statusRowLength;
        private IEnumerable<BusinessPartner> _cedents;
        private string _criteria;
        private bool _isSearching;
        private int _cedentCount;

        public int CedentCount
        {
            get => _cedentCount;
            set
            {
                _cedentCount = value; 
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("CedentCountLabel");
            }
        }

        public string CedentCountLabel => CedentCount > 0 ? $"{CedentCount:N0} {CedentMatchesString}" : string.Empty;
        
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

        public bool ShowMyCedents
        {
            get => _showMyCedents;
            set
            {
                _showMyCedents = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength CriteriaRowLength
        {
            get => _criteriaRowLength;
            set
            {
                _criteriaRowLength = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength StatusRowLength
        {
            get => _statusRowLength;
            set
            {
                _statusRowLength = value;
                NotifyPropertyChanged();
            }
        }

        public IEnumerable<BusinessPartner> Cedents
        {
            get => _cedents;
            set
            {
                _cedents = value;
                NotifyPropertyChanged();
            }
        }

        public string Criteria
        {
            get => _criteria;
            set
            {
                _criteria = value; 
                NotifyPropertyChanged();
                QueryCedent();
            }
        }

        public void QueryCedent()
        {
            try
            {
                IsSearching = true;
                Cedents = null;

                var criteria = Criteria;
                if (string.IsNullOrEmpty(criteria))
                {
                    SetScreenToNoLongerSearching();
                    CedentCount = 0;
                    return;
                }

                if (int.TryParse(criteria, out _)) criteria = criteria.PadLeft(BusinessPartnerCodeLength, '0');
                
                StatusRowLength = StatusRowLengthWhenVisible;
                StatusMessage = "Search in progress ...";
                CedentCount = 0;

                var cancellationTokenSource = new CancellationTokenSource();

                var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
                var task = new Task<IEnumerable<BusinessPartner>>(() => GetMatchingItems(criteria));
                task.ContinueWith(task1 =>
                {
                    if (task1.IsFaulted)
                    {
                        if (task1.Exception?.InnerException != null)
                        {
                            StatusMessage = task1.Exception.InnerException.Message;
                        }
                    }

                    Cedents = task.Result.OrderBy(x => x.NameAndLocation).ToList();
                    if (!Cedents.Any())
                    {
                        StatusMessage = "No matches found";
                    }
                    else
                    {
                        SetScreenToNoLongerSearching();
                        CedentCount = Cedents.Count();
                    }

                }, cancellationTokenSource.Token, TaskContinuationOptions.None, scheduler);

                task.Start();
            }
            catch (Exception ex)
            {
                MessageHelper.Show($"Cedent finder failed: {ex.Message}", MessageType.Stop);
                IsSearching = false;
            }
            
        }

        private void SetScreenToNoLongerSearching()
        {
            IsSearching = false;
            StatusMessage = string.Empty;
            StatusRowLength = NoLength;
        }

        private static IEnumerable<BusinessPartner> GetMatchingItems(string cedent)
        {
            var cedents = CedentSelector.CedentFinder.Find(cedent).ToList();
            cedents.ForEach(x => x.Id = x.ClippedId);
            return cedents;
        }
    }
}
