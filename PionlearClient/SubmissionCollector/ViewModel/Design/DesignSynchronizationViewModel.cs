namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignSynchronizationViewModel : BaseSynchronizationViewModel
    {
        public DesignSynchronizationViewModel()
        {
            LoadingRowPixels = LoadingRowHeight;
            ValidationRowPixels = ValidationRowHeight;
            ErrorRowPixels = ErrorRowHeight;
            BodyRowPixels = BodyRowHeight;
            LegendRowPixels = LegendRowHeight;
            StatusMessage = "Loading ...";
            ValidationMessage = "Can't process request: validation failed";
            ErrorMessage = "Can't process request: error";
            ShowZoom = true;

            var packageView = new SynchronizationView
            {
                Name = "RLI",
                SynchronizationCode = SynchronizationCode.InSynchronization,
            };

            var segmentView = new SynchronizationView
            {
                Name = "My Auto",
                SynchronizationCode = SynchronizationCode.InSynchronization,
            };

            segmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.InSynchronization,
                Name = "Hazard Profile AL - Comm"
            });
            segmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.InSynchronization,
                Name = "Hazard Profile GL - Prem Ops"
            });
            segmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.NotInSynchronization,
                Name = "Policy Profile AL - Comm, GL - Prem Ops"
            });
            segmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.New,
                Name = "State Profile AL - Comm, GL - Prem Ops"
            });
            segmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.Deleted,
                Name = "Exposures AL - Comm, GL - Prem Ops"
            });
            segmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.InSynchronization,
                Name = "Aggregate Loss AL - Comm, GL - Prem Ops"
            });
            segmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.InSynchronization,
                Name = "Individual Loss AL - Comm, GL - Prem Ops"
            });


            var anotherSegmentView = new SynchronizationView
            {
                Name = "My GL",
                SynchronizationCode = SynchronizationCode.InSynchronization
            };

            anotherSegmentView.ChildViews.Add(new SynchronizationView
            {
                SynchronizationCode = SynchronizationCode.InSynchronization,
                Name = "Hazard Profile AL - Comm"
            });

            packageView.ChildViews.Add(segmentView);
            packageView.ChildViews.Add(anotherSegmentView);
            PackageSynchronizationViews.Add(packageView);
        }
    }
}
