using System.Collections.ObjectModel;
using PionlearClient;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignUmbrellaTypeAllocatorViewModel : BaseUmbrellaTypeAllocatorViewModel
    {
        public DesignUmbrellaTypeAllocatorViewModel()
        {
            Title = $"Select {BexConstants.UmbrellaTypeName}s";
            
            UmbrellaItems = new ObservableCollection<UmbrellaTypeViewModelPlus>
            {
                new UmbrellaTypeViewModelPlus {IsSelected = true, UmbrellaTypeName = "Umbrella Type 1"},
                new UmbrellaTypeViewModelPlus {IsSelected = false, UmbrellaTypeName = "Umbrella Type 2"},
                new UmbrellaTypeViewModelPlus {IsSelected = false, UmbrellaTypeName = "Umbrella Type 3"}
            };
        }
    }
}
