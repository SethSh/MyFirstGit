using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignWorkersCompClassCodeViewModel: BaseWorkersCompClassCodeViewModel
    {
        public DesignWorkersCompClassCodeViewModel()
        {
            StateAbbreviations = new List<string> { "AL", "AK" };
            StateAbbreviationSelected = "AL";

            var state = new StateViewModel { Id = 25, Name = "Seth", DisplayOrder = 0, Abbreviation = "AL" };
            var coll = new List<WorkersCompClassCodeViewItem>
            {
                new WorkersCompClassCodeViewItem {StateClassCode = 1234, HazardGroupName = "HG A", StateDescription = "Hello World"},
                new WorkersCompClassCodeViewItem {StateClassCode = 5678, HazardGroupName = "HG B", StateDescription = "Hello Steve"},
            };

            var view = new WorkersCompClassCodeView { State = state, ClassCodeModels = coll };
            WorkersCompClassCodeViews.Add(view);


            var state2 = new StateViewModel { Id = 250, Name = "Jeff", DisplayOrder = 1, Abbreviation = "NJ" };
            var coll2 = new List<WorkersCompClassCodeViewItem>
            {
                new WorkersCompClassCodeViewItem {StateClassCode = 12, HazardGroupName = "HG A", StateDescription = "Hello World 2"},
                new WorkersCompClassCodeViewItem {StateClassCode = 13, HazardGroupName = "HG B", StateDescription = "Hello Steve 2"},
            };

            var view2 = new WorkersCompClassCodeView { State = state2, ClassCodeModels = coll2 };
            WorkersCompClassCodeViews.Add(view2);
            WorkersCompClassCodeView = WorkersCompClassCodeViews.First();
        }

        
    }
}
