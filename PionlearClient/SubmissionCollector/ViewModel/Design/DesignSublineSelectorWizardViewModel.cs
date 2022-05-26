using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using PionlearClient;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.Properties;

namespace SubmissionCollector.ViewModel.Design
{
    public class DesignSublineSelectorWizardViewModel : BaseSublineSelectorWizardViewModel
    {
        public DesignSublineSelectorWizardViewModel()
        {
            Log = "Feedback to help understand wizard behaviors";
            AreSublinesCustomizable = true;
            IsUserPreferenceAllowSublinesCustomization = true;
            SelectedSublinesRowSpan = GetSelectedSublineRowSpan();
            PolicyRowHeight = new GridLength(1, GridUnitType.Star);
            StateRowHeight = new GridLength(1, GridUnitType.Star); 
            OccupancyTypeRowHeight = new GridLength(1, GridUnitType.Star);
            TotalInsuredValueRowHeight = new GridLength(1, GridUnitType.Star);
            ExposureRowHeight = new GridLength(1, GridUnitType.Star);
            AggregateLossRowHeight = new GridLength(1, GridUnitType.Star);
            IndividualLossRowHeight = new GridLength(1, GridUnitType.Star);
            RateChangeRowHeight= new GridLength(1, GridUnitType.Star);

            var s1 = new FakeSubline {ImageSource = Resources.AngryBird.ToBitmapSource()};
            var s2 = new FakeSubline {ImageSource = Resources.AngryBird.ToBitmapSource()};
            var sublines = new ObservableCollection<ISubline>(new List<ISubline> {s1, s2});
            SegmentSublines = sublines;

            AvailableSublines = sublines;
            ComponentViews = new List<ComponentView>();

            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                ComponentViews.Add(CreateComponentView(i, ComponentViewType.Policy, BexConstants.PolicyProfileName));
                ComponentViews.Add(CreateComponentView(i, ComponentViewType.State, BexConstants.StateProfileName));
                ComponentViews.Add(CreateComponentView(i, ComponentViewType.ExposureSet, BexConstants.ExposureSetName));
                ComponentViews.Add(CreateComponentView(i, ComponentViewType.AggregateLossSet, BexConstants.AggregateLossSetName));
                ComponentViews.Add(CreateComponentView(i, ComponentViewType.IndividualLossSet, BexConstants.IndividualLossSetName));
                ComponentViews.Add(CreateComponentView(i, ComponentViewType.RateChangeSet, BexConstants.RateChangeSetName));
                
            }

            foreach (var view in ComponentViews)
            {
                view.RowIndex = ComponentViewRowIndices[view.Type];
                view.IsNew = true;
                view.HasData = true;
                view.IsVisible = true;
                view.IsLocked = true;
                view.IsMoveLeftButtonVisible = true;
                view.IsMoveRightButtonVisible = true;
                view.IsAddButtonVisible = true;
                view.IsDeleteButtonVisible = true;
                view.Sublines = new ObservableCollection<ISubline>(sublines);
            }
        }

        private ComponentView CreateComponentView(int row, ComponentViewType componentViewType, string name)
        {
            return new ComponentView(componentViewType)
            {
                Id = row,
                DisplayOrder = row,
                Name = name,
            };
        }
    }
}