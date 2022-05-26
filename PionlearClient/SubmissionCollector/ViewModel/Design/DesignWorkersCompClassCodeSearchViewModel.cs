namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignWorkersCompClassCodeSearchViewModel: BaseWorkersCompClassCodeSearchViewModel
    {
        public DesignWorkersCompClassCodeSearchViewModel()
        {
            FilteredClassCodeViewItems.Add(new WorkersCompClassCodeQueryViewItem
            {
                StateAbbreviation = "AL",
                StateClassCode = 1234,
                HazardGroupName = "HG A",
                StateDescription = "Hello World"
            });

            FilteredClassCodeViewItems.Add(new WorkersCompClassCodeQueryViewItem
            {
                StateAbbreviation = "AL",
                StateClassCode = 5678,
                HazardGroupName = "HG B",
                StateDescription = "Hello Steve"
            });

            FilteredClassCodeViewItems.Add(new WorkersCompClassCodeQueryViewItem
            {
                StateAbbreviation = "NJ",
                StateClassCode = 12,
                HazardGroupName = "HG A",
                StateDescription = "Hello World 2"
            });

            FilteredClassCodeViewItems.Add(new WorkersCompClassCodeQueryViewItem
            {
                StateAbbreviation = "NJ",
                StateClassCode = 13,
                HazardGroupName = "HG B",
                StateDescription = "Hello Steve 2"
            });

            SearchCriteria = "World";
            StatusMessage = "Search in progress ...";
        }
    }
}
