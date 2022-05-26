using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using PionlearClient.KeyDataFolder;

namespace SubmissionCollector.ViewModel
{
    public class UnderwriterSelectorViewModel : BaseUnderwriterSelectorViewModel
    {
        private string _criteria;
        private bool _showMyUnderwriters;

        public UnderwriterSelectorViewModel()
        {
            Underwriters = UnderwritersFromKeyData.UnderwriterReferenceData;
            Criteria = string.Empty;
            UnderwriterCount = 0;
            
            var up = UserPreferences.ReadFromFile();
            ShowMyUnderwriters = up.ShowMyUnderwriters;
        }

        public sealed override string Criteria
        {
            get => _criteria;
            set
            {
                _criteria = value;
                NotifyPropertyChanged();
                FilterUnderwriters();
            }
        }

        public sealed override bool ShowMyUnderwriters
        {
            get => _showMyUnderwriters;
            set
            {
                _showMyUnderwriters = value;
                NotifyPropertyChanged();
                FilterUnderwriters();
            }
        }
    }

    public abstract class BaseUnderwriterSelectorViewModel : ViewModelBase
    {
        private int _underwriterCount;
        private string _name;
        private string _code;
        private IList<Underwriter> _filteredUnderwriters;
        private GridLength _criteriaRowLength;
        public readonly GridLength NoLength = new GridLength(0);
        public readonly GridLength CriteriaRowLengthWhenVisible = new GridLength(40);

        public GridLength CriteriaRowLength
        {
            get => _criteriaRowLength;
            set
            {
                _criteriaRowLength = value;
                NotifyPropertyChanged();
            }
        }
        
        public int UnderwriterCount
        {
            get => _underwriterCount;
            set
            {
                _underwriterCount = value; 
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("UnderwriterCountLabel");
            }
        }

        public string UnderwriterCountLabel => UnderwriterCount > 0 ? $"{UnderwriterCount:N0} underwriter match(es)" : string.Empty;

        public string Name
        {
            get => _name;
            set
            {
                _name = value; 
                NotifyPropertyChanged();
            }
        }

        public string Code
        {
            get => _code;
            set
            {
                _code = value; 
                NotifyPropertyChanged();
            }
        }

        public IList<Underwriter> Underwriters { get; set; }

        public virtual string Criteria { get; set; }
        public virtual bool ShowMyUnderwriters{ get; set; }

        public IList<Underwriter> FilteredUnderwriters
        {
            get => _filteredUnderwriters;
            set
            {
                _filteredUnderwriters = value; 
                NotifyPropertyChanged();
            }
        }

        public void FilterUnderwriters()
        {
            if (ShowMyUnderwriters)
            {
                CriteriaRowLength = NoLength;
                UnderwriterCount = 0;

                var up = UserPreferences.ReadFromFile();
                FilteredUnderwriters = up.MyUnderwriters?.OrderBy(x => x.Name).ToList();
            }
            else
            {
                CriteriaRowLength = CriteriaRowLengthWhenVisible;
                if (string.IsNullOrEmpty(Criteria))
                {
                    FilteredUnderwriters = Underwriters;
                    UnderwriterCount = 0;
                }
                else
                {
                    FilteredUnderwriters = Underwriters.Where(u => u.Name.IndexOf(Criteria, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                    UnderwriterCount = FilteredUnderwriters.Any() ? FilteredUnderwriters.Count : 0;
                }
            }
        }
    }
}
