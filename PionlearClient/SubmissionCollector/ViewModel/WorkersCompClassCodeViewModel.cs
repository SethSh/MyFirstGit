using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using PionlearClient.BexReferenceData;

namespace SubmissionCollector.ViewModel
{
    internal class WorkersCompClassCodeViewModel : BaseWorkersCompClassCodeViewModel
    {
        public WorkersCompClassCodeViewModel()
        {
            var stateClassCodeSets = WorkersCompClassCodesAndHazardsFromBex.StateClassCodes.ToList();
            StateAbbreviations = stateClassCodeSets.Select(codeSet => codeSet.State.Abbreviation).OrderBy(item => item);
            StateAbbreviationSelected = StateAbbreviations.First();

            var hazardGroups = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.ToDictionary(key => key.Id);
            
            foreach (var codeSet in stateClassCodeSets)
            {
                var state = codeSet.State;
                var classCodes = codeSet.ClassCodes.OrderBy(classCode => classCode.StateClassCode).Select(classCode => new WorkersCompClassCodeViewItem
                {
                    StateClassCode = classCode.StateClassCode,
                    HazardGroupName = classCode.HazardGroupId.HasValue ? hazardGroups[classCode.HazardGroupId.Value].Name : string.Empty,
                    StateDescription = classCode.StateDescription
                });

                var view = new WorkersCompClassCodeView { State = state, ClassCodeModels = classCodes };
                WorkersCompClassCodeViews.Add(view);
            }
        }
    }

    internal abstract class BaseWorkersCompClassCodeViewModel : ViewModelBase
    {
        private string _stateAbbreviationSelected;
        private WorkersCompClassCodeView _workersCompClassCodeView;

        public ObservableCollection<WorkersCompClassCodeView> WorkersCompClassCodeViews { get; set; }

        public WorkersCompClassCodeView WorkersCompClassCodeView
        {
            get => _workersCompClassCodeView;
            set
            {
                _workersCompClassCodeView = value;
                NotifyPropertyChanged();
            }
        }

        public IEnumerable<string> StateAbbreviations { get; set; }

        public string StateAbbreviationSelected
        {
            get => _stateAbbreviationSelected;
            set
            {
                _stateAbbreviationSelected = value;
                NotifyPropertyChanged();
            }
        }

        protected BaseWorkersCompClassCodeViewModel()
        {
            WorkersCompClassCodeViews = new ObservableCollection<WorkersCompClassCodeView>();
        }
    }

    internal class WorkersCompClassCodeView : ViewModelBase
    {
        public StateViewModel State { get; set; }
        public IEnumerable<WorkersCompClassCodeViewItem> ClassCodeModels { get; set; }
        
        public WorkersCompClassCodeView()
        {
            State = new StateViewModel();
            ClassCodeModels = new List<WorkersCompClassCodeViewItem>();
        }
    }

    internal class WorkersCompClassCodeViewItem
    {
        public string StateClassCodeAsString => StateClassCode.ToString("0000");
        public long StateClassCode { get; set; }
        public string HazardGroupName { get; set; }
        public string StateDescription { get; set; }
    }
    
}