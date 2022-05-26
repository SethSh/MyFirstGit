using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using PionlearClient;
using SubmissionCollector.Models.Comparers;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.View;

namespace SubmissionCollector.ViewModel
{
    public interface ISublineSelectorWizardViewModel
    {
        ObservableCollection<ISubline> SegmentSublines { get; set; }
        
        IEnumerable<ComponentView> AggregateLossSets { get; }
        IEnumerable<ComponentView> IndividualLossSets { get; }
        IEnumerable<ComponentView> RateChangeSets { get; }
        IEnumerable<ComponentView> ExposureSets { get; }
        
        IEnumerable<ComponentView> PolicyProfiles { get; }
        IEnumerable<ComponentView> StateProfiles { get; }
    }

    public class SublineWizardViewModel : BaseSublineSelectorWizardViewModel
    {
        public SublineWizardViewModel(ISegment segment)
        {
            Segment = segment;
            Log = string.Empty;
            AutomaticallySyncSublines = true;
            ShowMoveLeftAndMoveRightButtons = true;

            var allSublines = GetAllSublines(segment.Id).ToList();
            var anySublines = segment.Count > 0;
            var segmentSelectedSublines = anySublines
                ? segment.OrderBy(subline => subline.ShortNameWithLob).ToList()
                : new List<ISubline>();

            AvailableSublines = new ObservableCollection<ISubline>(allSublines.Except(segmentSelectedSublines, new SublineComparer()));
            SegmentSublines = new ObservableCollection<ISubline>(segmentSelectedSublines);

            InitializeMaxIndices();

            ComponentViews = new List<ComponentView>();
            CreatePolicyViews(segment);
            CreateStateViews(segment);
            
            CreateExposureSetViews(segment);
            CreateAggregateLossSetViews(segment);
            CreateIndividualLossSetViews(segment);
            CreateRateChangeSetViews(segment);
            
            ComponentViews.ForEach(view => view.RowIndex = ComponentViewRowIndices[view.Type]);

            SetSublineCustomization();
            SelectedSublinesRowSpan = GetSelectedSublineRowSpan();
        }

        internal void InitializeMaxIndices()
        {
            MaximumIndices = new Dictionary<ComponentViewType, int?>
            {
                {ComponentViewType.Policy, Segment.Any(x => x.HasPolicyProfile)  ? Segment.PolicyProfiles.Max(x => x.ComponentId) : new int?()},
                {ComponentViewType.State, Segment.Any(x => x.HasStateProfile) ? Segment.StateProfiles.Max(x => x.ComponentId) : new int?()},
                {ComponentViewType.ExposureSet, Segment.Count > 0 ? Segment.ExposureSets.Max(x => x.ComponentId) : new int?()},
                {ComponentViewType.AggregateLossSet, Segment.Count > 0 ? Segment.AggregateLossSets.Max(x => x.ComponentId) : new int?()},
                {ComponentViewType.IndividualLossSet, Segment.Count > 0 ? Segment.IndividualLossSets.Max(x => x.ComponentId) : new int?()},
            };

            if (Segment.RateChangeSets.Any())
            {
                MaximumIndices.Add(ComponentViewType.RateChangeSet,
                    Segment.Count > 0 ? Segment.RateChangeSets.Max(x => x.ComponentId) : new int?());
            }
            else
            {
                MaximumIndices.Add(ComponentViewType.RateChangeSet, new int?());
            }

        }

        internal void SetSublineCustomization()
        {
            var up = UserPreferences.ReadFromFile();
            IsUserPreferenceAllowSublinesCustomization = up.AllowSublineFlexibility;

            if (IsUserPreferenceAllowSublinesCustomization)
            {
                AreSublinesCustomizable = true;
                return;
            }
            
            AreSublinesCustomizable = GetComponentCountByType().Values.Max() > 1;
        }

        private void CreatePolicyViews(ISegment segment)
        {
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var view = new ComponentView(ComponentViewType.Policy) { DisplayOrder = i };
                var profile = segment.PolicyProfiles.SingleOrDefault(x => x.IntraDisplayOrder == i);

                if (profile != null)
                {
                    CreateView(view, profile);
                    view.UmbrellaType = profile.UmbrellaType;
                }
                else
                {
                    CreateEmptyView(view);
                }
                ComponentViews.Add(view);
            }
        }

        private void CreateStateViews(ISegment segment)
        {
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var view = new ComponentView(ComponentViewType.State) { DisplayOrder = i };
                var profile = segment.StateProfiles.SingleOrDefault(x => x.IntraDisplayOrder == i);

                if (profile != null)
                {
                    CreateView(view, profile);
                }
                else
                {
                    CreateEmptyView(view);
                }
                ComponentViews.Add(view);
            }
        }


        private void CreateExposureSetViews(ISegment segment)
        {
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var view = new ComponentView(ComponentViewType.ExposureSet) { DisplayOrder = i };
                var exposureSet = segment.ExposureSets.SingleOrDefault(x => x.IntraDisplayOrder == i);

                if (exposureSet != null)
                {
                    CreateView(view, exposureSet);
                }
                else
                {
                    CreateEmptyView(view);
                }
                ComponentViews.Add(view);
            }
        }

        private void CreateAggregateLossSetViews(ISegment segment)
        {
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var view = new ComponentView(ComponentViewType.AggregateLossSet) {DisplayOrder = i};
                var lossSet = segment.AggregateLossSets.SingleOrDefault(x => x.IntraDisplayOrder == i);

                if (lossSet != null)
                {
                    CreateView(view, lossSet);
                }
                else
                {
                    CreateEmptyView(view);
                }
                ComponentViews.Add(view);
            }
        }

        private void CreateIndividualLossSetViews(ISegment segment)
        {
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var view = new ComponentView(ComponentViewType.IndividualLossSet) {DisplayOrder = i};
                var lossSet = segment.IndividualLossSets.SingleOrDefault(x => x.IntraDisplayOrder == i);

                if (lossSet != null)
                {
                    CreateView(view, lossSet);
                }
                else
                {
                    CreateEmptyView(view);
                }
                ComponentViews.Add(view);
            }
        }

        private void CreateRateChangeSetViews(ISegment segment)
        {
            var type = ComponentViewType.RateChangeSet;
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var view = new ComponentView(type) { DisplayOrder = i };
                var rateChangeSet = segment.RateChangeSets.SingleOrDefault(x => x.IntraDisplayOrder == i);

                if (rateChangeSet != null)
                {
                    CreateView(view, rateChangeSet);
                }
                else
                {
                    CreateEmptyView(view);
                }
                ComponentViews.Add(view);
            }

            var isSegmentExisting = segment.Any();
            var singleView = ComponentViews.Single(v => v.Type == ComponentViewType.RateChangeSet && v.DisplayOrder == 0);
            if (isSegmentExisting && !segment.RateChangeSets.Any()) RetroFitRateChangeSet(segment, singleView);
        }

        private void RetroFitRateChangeSet(ISegment segment, ComponentView view)
        {
            view.Id = 0;
            view.Guid = Guid.NewGuid();
            view.Name = BexConstants.RateChangeSetName;
            view.IsNew = true;
            view.HasData = false;
            view.IsVisible = true;
            view.BorderThickness = DefaultBorderThickness;
            view.IsLocked = false;
            foreach (var subline in segment)
            {
                view.Sublines.Add(subline);
            }

            var indexPlusOne = MaximumIndices[view.Type].AddOne();
            MaximumIndices[view.Type] = indexPlusOne;
        }


        private static void CreateView(ComponentView view, IExcelComponent excelComponent)
        {
            view.Id = excelComponent.ComponentId;
            view.Guid = excelComponent.Guid;
            view.Name = excelComponent.Name;
            view.IsNew = false;
            view.HasData = excelComponent.CommonExcelMatrix.HasData;
            view.IsVisible = true;
            view.BorderThickness = DefaultBorderThickness;
            view.IsLocked = excelComponent.SourceId.HasValue;
            view.Sublines = new ObservableCollection<ISubline>( ((MultipleOccurrenceSegmentExcelMatrix)excelComponent.CommonExcelMatrix).OrderBy(x => x.ShortNameWithLob));
        }
        

        private static void CreateEmptyView(ComponentView view)
        {
            view.IsVisible = false;
            view.IsLocked = false;
            view.BorderThickness = 0;
            view.Sublines = new ObservableCollection<ISubline>();
        }

        private static IEnumerable<ISubline> GetAllSublines(int segmentId)
        {
            var sublines = typeof(BaseSubline)
                .Assembly.GetTypes()
                .Where(t => t.IsSubclassOf(typeof(BaseSubline)) && !t.IsAbstract)
                .Select(t => (BaseSubline) Activator.CreateInstance(t))
                .OrderBy(x => x.ShortNameWithLob).ToList();

            sublines.ForEach(sub => sub.SegmentId = segmentId);
            return sublines;
        }
    }

    public abstract class BaseSublineSelectorWizardViewModel : ViewModelBase, ISublineSelectorWizardViewModel
    {
        private string _log;
        private bool _showMoveLeftAndMoveRightButtons;
        internal const int DefaultBorderThickness = 1;
        public ISegment Segment;
        private int _selectedSublinesRowSpan;
        private bool _areSublinesCustomizable;
        private bool _isUserPreferenceAllowSublinesCustomization;
        private string _okButtonToolTip;
        private bool _okButtonEnabled;
        private GridLength _policyRowHeight;
        private GridLength _stateRowHeight;
        private GridLength _occupancyTypeRowHeight; 
        private GridLength _totalInsuredValueRowHeight;
        private GridLength _exposureRowHeight;
        private GridLength _aggregateLossRowHeight;
        private GridLength _individualLossRowHeight;
        private GridLength _rateChangeRowHeight;

        protected BaseSublineSelectorWizardViewModel()
        {
            ComponentViewNamePrefixes = new Dictionary<ComponentViewType, string>
            {
                {ComponentViewType.Policy, BexConstants.PolicyProfileName},
                {ComponentViewType.State, BexConstants.StateProfileName},
                {ComponentViewType.ExposureSet, BexConstants.ExposureSetName},
                {ComponentViewType.AggregateLossSet, BexConstants.AggregateLossSetName},
                {ComponentViewType.IndividualLossSet, BexConstants.IndividualLossSetName},
                {ComponentViewType.RateChangeSet, BexConstants.RateChangeSetName},
            };

            var counter = 0;
            ComponentViewRowIndices = new Dictionary<ComponentViewType, int>
            {
                {ComponentViewType.Policy, counter++},
                {ComponentViewType.State, counter++},
                {ComponentViewType.ExposureSet, counter++},
                {ComponentViewType.AggregateLossSet, counter++},
                {ComponentViewType.IndividualLossSet, counter++},
                {ComponentViewType.RateChangeSet, counter},
            };
        }

        public Dictionary<ComponentViewType, int> ComponentViewRowIndices { get; set; }

        public GridLength PolicyRowHeight
        {
            get => _policyRowHeight;
            set
            {
                _policyRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength StateRowHeight
        {
            get => _stateRowHeight;
            set
            {
                _stateRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength OccupancyTypeRowHeight
        {
            get => _occupancyTypeRowHeight;
            set
            {
                _occupancyTypeRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        
        public GridLength TotalInsuredValueRowHeight
        {
            get => _totalInsuredValueRowHeight;
            set
            {
                _totalInsuredValueRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength ExposureRowHeight
        {
            get => _exposureRowHeight;
            set
            {
                _exposureRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength AggregateLossRowHeight
        {
            get => _aggregateLossRowHeight;
            set
            {
                _aggregateLossRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength IndividualLossRowHeight
        {
            get => _individualLossRowHeight;
            set
            {
                _individualLossRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength RateChangeRowHeight
        {
            get => _rateChangeRowHeight;
            set
            {
                _rateChangeRowHeight = value;
                NotifyPropertyChanged();
            }
        }

        public int SelectedSublinesRowSpan
        {
            get => _selectedSublinesRowSpan;
            set
            {
                _selectedSublinesRowSpan = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsUserPreferenceAllowSublinesCustomization
        {
            get => _isUserPreferenceAllowSublinesCustomization;
            set
            {
                _isUserPreferenceAllowSublinesCustomization = value;
                NotifyPropertyChanged();
            }
        }

        public bool AreSublinesCustomizable
        {
            get => _areSublinesCustomizable;
            set
            {
                _areSublinesCustomizable = value;
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("AreSublinesRigid");
            }
        }

        public bool AreSublinesRigid => !AreSublinesCustomizable;

        public string Log
        {
            get => _log;
            set
            {
                _log = value;
                NotifyPropertyChanged();
            }
        }
        
        public bool AutomaticallySyncSublines { get; set; }

        public bool ShowMoveLeftAndMoveRightButtons
        {
            get => _showMoveLeftAndMoveRightButtons;
            set
            {
                _showMoveLeftAndMoveRightButtons = value; 
                NotifyPropertyChanged();
            }
        }
        
        public ObservableCollection<ISubline> AvailableSublines { get; set; }
        public ObservableCollection<ISubline> SegmentSublines { get; set; }

        public int MaxDisplayOrder => BexConstants.MaximumNumberOfDataComponents - 1;
        public List<ComponentView> ComponentViews { get; set; }

        #region PolicyProfiles
        public IEnumerable<ComponentView> PolicyProfiles => ComponentViews.Where(x => x.Type == ComponentViewType.Policy).OrderBy(x => x.DisplayOrder);
        
        public ComponentView PolicyProfile1
        {
            get { return PolicyProfiles.Single(x => x.DisplayOrder == 0); }
        }

        public ComponentView PolicyProfile2
        {
            get { return PolicyProfiles.Single(x => x.DisplayOrder == 1); }
        }

        public ComponentView PolicyProfile3
        {
            get { return PolicyProfiles.Single(x => x.DisplayOrder == 2); }
        }

        public ComponentView PolicyProfile4
        {
            get { return PolicyProfiles.Single(x => x.DisplayOrder == 3); }
        }

        public ComponentView PolicyProfile5
        {
            get { return PolicyProfiles.Single(x => x.DisplayOrder == 4); }
        }

        public ComponentView PolicyProfile6
        {
            get { return PolicyProfiles.Single(x => x.DisplayOrder == 5); }
        }
        #endregion

        #region StateProfile
        public IEnumerable<ComponentView> StateProfiles => ComponentViews.Where(x => x.Type == ComponentViewType.State).OrderBy(x => x.DisplayOrder);

        public ComponentView StateProfile1
        {
            get { return StateProfiles.Single(x => x.DisplayOrder == 0); }
        }

        public ComponentView StateProfile2
        {
            get { return StateProfiles.Single(x => x.DisplayOrder == 1); }
        }

        public ComponentView StateProfile3
        {
            get { return StateProfiles.Single(x => x.DisplayOrder == 2); }
        }

        public ComponentView StateProfile4
        {
            get { return StateProfiles.Single(x => x.DisplayOrder == 3); }
        }

        public ComponentView StateProfile5
        {
            get { return StateProfiles.Single(x => x.DisplayOrder == 4); }
        }

        public ComponentView StateProfile6
        {
            get { return StateProfiles.Single(x => x.DisplayOrder == 5); }
        }
        #endregion

        
        #region Exposure Sets
        public IEnumerable<ComponentView> ExposureSets => ComponentViews.Where(x => x.Type == ComponentViewType.ExposureSet).OrderBy(x => x.DisplayOrder);

        public ComponentView ExposureSet1
        {
            get { return ExposureSets.Single(x => x.DisplayOrder == 0); }
        }

        public ComponentView ExposureSet2
        {
            get { return ExposureSets.Single(x => x.DisplayOrder == 1); }
        }

        public ComponentView ExposureSet3
        {
            get { return ExposureSets.Single(x => x.DisplayOrder == 2); }
        }

        public ComponentView ExposureSet4
        {
            get { return ExposureSets.Single(x => x.DisplayOrder == 3); }
        }

        public ComponentView ExposureSet5
        {
            get { return ExposureSets.Single(x => x.DisplayOrder == 4); }
        }

        public ComponentView ExposureSet6
        {
            get { return ExposureSets.Single(x => x.DisplayOrder == 5); }
        }
        #endregion

        #region Aggregate Loss Sets
        public IEnumerable<ComponentView> AggregateLossSets => ComponentViews.Where(x => x.Type == ComponentViewType.AggregateLossSet).OrderBy(x => x.DisplayOrder);

        public ComponentView AggregateLossSet1
        {
            get { return AggregateLossSets.Single(x => x.DisplayOrder == 0); }
        }

        public ComponentView AggregateLossSet2
        {
            get { return AggregateLossSets.Single(x => x.DisplayOrder == 1); }
        }

        public ComponentView AggregateLossSet3
        {
            get { return AggregateLossSets.Single(x => x.DisplayOrder == 2); }
        }

        public ComponentView AggregateLossSet4
        {
            get { return AggregateLossSets.Single(x => x.DisplayOrder == 3); }
        }

        public ComponentView AggregateLossSet5
        {
            get { return AggregateLossSets.Single(x => x.DisplayOrder == 4); }
        }

        public ComponentView AggregateLossSet6
        {
            get { return AggregateLossSets.Single(x => x.DisplayOrder == 5); }
        }
        #endregion

        #region Individual Loss Sets
        public IEnumerable<ComponentView> IndividualLossSets => ComponentViews.Where(x => x.Type == ComponentViewType.IndividualLossSet).OrderBy(x => x.DisplayOrder);

        public ComponentView IndividualLossSet1 => IndividualLossSets.Single(x => x.DisplayOrder == 0);
        public ComponentView IndividualLossSet2 => IndividualLossSets.Single(x => x.DisplayOrder == 1);
        public ComponentView IndividualLossSet3 => IndividualLossSets.Single(x => x.DisplayOrder == 2);
        public ComponentView IndividualLossSet4 => IndividualLossSets.Single(x => x.DisplayOrder == 3);
        public ComponentView IndividualLossSet5 => IndividualLossSets.Single(x => x.DisplayOrder == 4);
        public ComponentView IndividualLossSet6 => IndividualLossSets.Single(x => x.DisplayOrder == 5);
        #endregion

        #region Rate Change Sets
        public IEnumerable<ComponentView> RateChangeSets => ComponentViews.Where(x => x.Type == ComponentViewType.RateChangeSet).OrderBy(x => x.DisplayOrder);

        public ComponentView RateChangeSet1 => RateChangeSets.Single(x => x.DisplayOrder == 0);
        public ComponentView RateChangeSet2 => RateChangeSets.Single(x => x.DisplayOrder == 1);
        public ComponentView RateChangeSet3 => RateChangeSets.Single(x => x.DisplayOrder == 2);
        public ComponentView RateChangeSet4 => RateChangeSets.Single(x => x.DisplayOrder == 3);
        public ComponentView RateChangeSet5 => RateChangeSets.Single(x => x.DisplayOrder == 4);
        public ComponentView RateChangeSet6 => RateChangeSets.Single(x => x.DisplayOrder == 5);
        #endregion

        public Dictionary<ComponentViewType, int?> MaximumIndices { get; set; }
        public Dictionary<ComponentViewType, string> ComponentViewNamePrefixes { get; set; }

        public bool GetIsMoveLeftButtonVisible(int sortOrder)
        {
            return sortOrder > 0;
        }

        public bool GetIsMoveRightButtonVisible(ComponentViewType type, int sortOrder)
        {
            return sortOrder < ComponentViews.Where(x => x.Type == type && x.IsVisible).Max(x => x.DisplayOrder);
        }

        internal int GetSelectedSublineRowSpan()
        {
            return AreSublinesCustomizable ? 1 : 3;
        }

        public Dictionary<ComponentViewType, int> GetComponentCountByType()
        {
            return new Dictionary<ComponentViewType, int>
            {
                {ComponentViewType.Policy, GetVisibleCount(PolicyProfiles)},
                {ComponentViewType.State, GetVisibleCount(StateProfiles)},
                {ComponentViewType.ExposureSet, GetVisibleCount(ExposureSets)},
                {ComponentViewType.AggregateLossSet, GetVisibleCount(AggregateLossSets)},
                {ComponentViewType.IndividualLossSet, GetVisibleCount(IndividualLossSets)},
                {ComponentViewType.RateChangeSet, GetVisibleCount(RateChangeSets)},
            };
        }

        public bool OkButtonToolTipVisibility => !string.IsNullOrEmpty(OkButtonToolTip);

        public string OkButtonToolTip
        {
            get => _okButtonToolTip;
            set
            {
                _okButtonToolTip = value; 
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("OkButtonToolTipVisibility");
            }
        }

        public bool OkButtonEnabled
        {
            get => _okButtonEnabled;
            set
            {
                _okButtonEnabled = value;
                NotifyPropertyChanged();
            }
        }

        private static int GetVisibleCount(IEnumerable<ComponentView> list)
        {
            return list.Count(x => x.IsVisible);
        }
    }

    public class ComponentView : ViewModelBase
    {
        public bool ShowLock => IsVisible && IsLocked;
        private bool _isVisible;
        private ObservableCollection<ISubline> _sublines;
        private string _name;
        private bool _isNew;
        private bool _hasData;
        private int _displayOrder;
        private bool _isMoveLeftButtonVisible;
        private bool _isMoveRightButtonVisible;
        private bool _isAddButtonVisible;
        private bool _isDeleteButtonVisible;
        private int _borderThickness;
        
        public ComponentViewType Type { get; set; }

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                _isVisible = value;
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("ShowLock");
            }
        }
        
        public ObservableCollection<ISubline> Sublines
        {
            get => _sublines;
            set
            {
                _sublines = value;
                NotifyPropertyChanged();
            }
        }

        public bool HasData
        {
            get => _hasData;
            set
            {
                _hasData = value; 
                NotifyPropertyChanged();
            }
        }

        public bool IsNew
        {
            get => _isNew;
            set
            {
                _isNew = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsLocked { get; set; }
        public bool IsEnabled => !IsLocked;
        public bool IsMoveLeftButtonVisible
        {
            get => _isMoveLeftButtonVisible;
            set
            {
                _isMoveLeftButtonVisible = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsAddButtonVisible
        {
            get => _isAddButtonVisible;
            set
            {
                _isAddButtonVisible = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsDeleteButtonVisible
        {
            get => _isDeleteButtonVisible;
            set
            {
                _isDeleteButtonVisible = value;
                NotifyPropertyChanged();
            }
        }
        
        public bool IsMoveRightButtonVisible
        {
            get => _isMoveRightButtonVisible;
            set
            {
                _isMoveRightButtonVisible = value; 
                NotifyPropertyChanged();
            }
        }

        public int? UmbrellaType { get; set; }

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }
        
        public Guid Guid { get; set; }

        public int Id { get; set; }

        public int RowIndex { get; set; }

        public int ColumnIndex => DisplayOrder;
        
        public int DisplayOrder
        {
            get => _displayOrder;
            set
            {
                _displayOrder = value;
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("ColumnIndex");
            }
        }

        public int BorderThickness
        {
            get => _borderThickness;
            set
            {
                _borderThickness = value; 
                NotifyPropertyChanged();
            }
        }

        public ComponentView(ComponentViewType type)
        {
            Type = type;
            BorderThickness = 1;
        }
    }

    public enum ComponentViewType
    {
        Policy,
        State,
        ExposureSet,
        AggregateLossSet,
        IndividualLossSet,
        RateChangeSet
    }

}
