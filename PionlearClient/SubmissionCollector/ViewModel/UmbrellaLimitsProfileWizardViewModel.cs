using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.ViewModel
{
    public class UmbrellaTypePolicyProfileWizardViewModel : BaseUmbrellaPolicyProfileWizardViewModel
    {
        public ISegment Segment { get; set; }

        public UmbrellaTypePolicyProfileWizardViewModel(ISegment segment)
        {
            Segment = segment;
            Title = $"{segment.Name} {BexConstants.Dash} {BexConstants.PolicyProfileName}s";

            segment.UmbrellaExcelMatrix.Validate();

            SegmentUmbrellaTypes = new List<UmbrellaTypeViewModel>();
            foreach (var alloc in segment.UmbrellaExcelMatrix.Allocations)
            {
                var id = Convert.ToInt32(alloc.Id);
                if (UmbrellaTypesFromBex.GetIsPersonal(id)) continue;
                
                var name = UmbrellaTypesFromBex.GetName(id);
                SegmentUmbrellaTypes.Add(new UmbrellaTypeViewModel {UmbrellaTypeCode = id, UmbrellaTypeName = name});
            }

            ComponentViews = new List<UmbrellaComponentView>();
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var umbrellaWizardView = CreateUmbrellaWizardView(i);
                ComponentViews.Add(umbrellaWizardView);
            }

            PolicyProfilesMaxIndex = segment.PolicyProfiles.Max(x => x.ComponentId);
        }

        private UmbrellaComponentView CreateUmbrellaWizardView(int displayOrder)
        {
            var policyProfile = Segment.PolicyProfiles.SingleOrDefault(x => x.IntraDisplayOrder == displayOrder);
            if (policyProfile != null)
            {
                var umbrellaWizardView = new UmbrellaComponentView
                {
                    Id = policyProfile.ComponentId,
                    DisplayOrder = policyProfile.IntraDisplayOrder,
                    Guid = policyProfile.Guid,
                    Name = BexConstants.PolicyProfileName,
                    FullName = policyProfile.Name,
                    IsVisible = true,
                    BorderThickness = DefaultBorderThickness,
                    IsLocked = policyProfile.SourceId.HasValue,
                    IsNew = false,
                    HasData = policyProfile.ExcelMatrix.HasData,
                    IsShaded = false,
                    Sublines = new ObservableCollection<ISubline>(policyProfile.ExcelMatrix),
                    UmbrellaTypes = CreateUmbrellaView(displayOrder)
                };

                return umbrellaWizardView;
            }

            var umbrellaWizardView2 = new UmbrellaComponentView
            {
                Id = displayOrder,
                BorderThickness = DefaultBorderThickness,
                DisplayOrder = displayOrder,
                IsVisible = false
            };
            return umbrellaWizardView2;
        }

        private ObservableCollection<UmbrellaTypeViewModel> CreateUmbrellaView(int displayOrder)
        {
            var profile = Segment.PolicyProfiles.SingleOrDefault(x => x.IntraDisplayOrder == displayOrder);
            if (profile == null) return null;

            return profile.ExcelMatrix.All(subline => subline.IsPersonal) ? CreatePersonalUmbrellaView() : CreateCommercialUmbrellaView(profile);
        }

        private ObservableCollection<UmbrellaTypeViewModel> CreateCommercialUmbrellaView(PolicyProfile profile)
        {
            if (profile.UmbrellaType != null)
            {
                var referenceItem = UmbrellaTypesFromBex.ReferenceData.Single(x => x.UmbrellaTypeCode == profile.UmbrellaType);
                var umbrellaTypeViewModel = new UmbrellaTypeViewModel
                {
                    UmbrellaTypeName = referenceItem.UmbrellaTypeName,
                    UmbrellaTypeCode = referenceItem.UmbrellaTypeCode
                };

                return new ObservableCollection<UmbrellaTypeViewModel> { umbrellaTypeViewModel };
            }

            var umbrellaTypes = new ObservableCollection<UmbrellaTypeViewModel>();
            Segment.UmbrellaExcelMatrix.Validate();
            foreach (var allocation in Segment.UmbrellaExcelMatrix.Allocations)
            {
                if (UmbrellaTypesFromBex.GetIsPersonal(Convert.ToInt32(allocation.Id))) continue;

                var name = UmbrellaTypesFromBex.GetName(allocation.Id);
                var umbrellaTypeWithAllocationViewModel = new UmbrellaTypeViewModel
                {
                    UmbrellaTypeName = name,
                    UmbrellaTypeCode = Convert.ToInt32(allocation.Id)
                };
                umbrellaTypes.Add(umbrellaTypeWithAllocationViewModel);
            }

            return umbrellaTypes;
        }

        private ObservableCollection<UmbrellaTypeViewModel> CreatePersonalUmbrellaView()
        {
            var viewModel = UmbrellaTypesFromBex.ReferenceData.Single(x => x.IsPersonal);
            var umbrellaTypeViewModel = new UmbrellaTypeViewModel
            {
                UmbrellaTypeName = viewModel.UmbrellaTypeName,
                UmbrellaTypeCode = viewModel.UmbrellaTypeCode
            };

            return new ObservableCollection<UmbrellaTypeViewModel> {umbrellaTypeViewModel};
        }
    }

    public abstract class BaseUmbrellaPolicyProfileWizardViewModel : ViewModelBase, IUmbrellaPolicyProfileWizardViewModel
    {
        public string Title { get; set; }
        public List<UmbrellaComponentView> ComponentViews { get; set; }
        
        public UmbrellaComponentView Component1 => ComponentViews.Single(x => x.DisplayOrder == 0);
        public UmbrellaComponentView Component2 => ComponentViews.Single(x => x.DisplayOrder == 1);
        public UmbrellaComponentView Component3 => ComponentViews.Single(x => x.DisplayOrder == 2);
        public UmbrellaComponentView Component4 => ComponentViews.Single(x => x.DisplayOrder == 3);
        public UmbrellaComponentView Component5 => ComponentViews.Single(x => x.DisplayOrder == 4);
        public UmbrellaComponentView Component6 => ComponentViews.Single(x => x.DisplayOrder == 5);
        
        public IList<UmbrellaTypeViewModel> SegmentUmbrellaTypes { get; protected set; }

        public int PolicyProfilesMaxIndex { get; set; }

        public int MaxDisplayOrder => BexConstants.MaximumNumberOfDataComponents - 1;

        internal const int DefaultBorderThickness = 1;

        public bool GetIsMoveLeftButtonVisible(int displayOrder)
        {
            return displayOrder > 0;
        }

        public bool GetIsMoveRightButtonVisible(int displayOrder)
        {
            return displayOrder < ComponentViews.Where(x=> x.IsVisible).Max(x => x.DisplayOrder);
        }
    }

    internal interface IUmbrellaPolicyProfileWizardViewModel
    {
        string Title { get; set; }
        
        List<UmbrellaComponentView> ComponentViews { get; set; }
        
        UmbrellaComponentView Component1 { get; }
        UmbrellaComponentView Component2 { get; }
        UmbrellaComponentView Component3 { get; }
        UmbrellaComponentView Component4 { get; }
        UmbrellaComponentView Component5 { get; }
        UmbrellaComponentView Component6 { get; }

        IList<UmbrellaTypeViewModel> SegmentUmbrellaTypes { get; }
    }

    public class UmbrellaComponentView : ViewModelBase
    {
        private ObservableCollection<UmbrellaTypeViewModel> _umbrellaTypes;
        private bool _isVisible;
        private bool _isLocked;
        private bool _isDivideButtonEnabled;
        private bool _isDivideButtonVisible;
        private bool _isMergeButtonVisible;
        private bool _isMergeButtonEnabled;
        private string _name;
        private ObservableCollection<ISubline> _sublines;
        private bool _isShaded;
        private string _divideToolTip;
        private bool _hasData;
        private bool _isNew;
        private int _borderThickness;
        private int _displayOrder;
        private bool _isMoveRightButtonVisible;
        private bool _isMoveLeftButtonVisible;
        
        public int Id { get; set; }

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

        public int ColumnIndex => DisplayOrder + 1;

        public string FullName { get; set; }

        public string Name
        {
            get => _name;
            set
            {
                _name = value; 
                NotifyPropertyChanged();
            }
        }

        public bool IsShaded
        {
            get => _isShaded;
            set
            {
                _isShaded = value;
                NotifyPropertyChanged();
            }
        }
        
        public ObservableCollection<UmbrellaTypeViewModel> UmbrellaTypes
        {
            get => _umbrellaTypes;
            set
            {
                _umbrellaTypes = value;
                NotifyPropertyChanged();
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

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                _isVisible = value;
                NotifyPropertyChanged();
            }
        }
        public bool IsLocked
        {
            get => _isLocked;
            set
            {
                _isLocked = value;
                NotifyPropertyChanged();
            }
        }
        public bool IsEnabled => !IsLocked;

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

        public int BorderThickness
        {
            get => _borderThickness;
            set
            {
                _borderThickness = value; 
                NotifyPropertyChanged();
            }
        }

        public Guid Guid { get; set; }
        public bool IsDivideButtonEnabled
        {
            get => _isDivideButtonEnabled;
            set
            {
                _isDivideButtonEnabled = value;
                NotifyPropertyChanged();
            }
        }
        public bool IsDivideButtonVisible 
        {
            get => _isDivideButtonVisible;
            set
            {
                _isDivideButtonVisible = value;
                NotifyPropertyChanged();
            }
        }

        public string DivideToolTip 
        {
            get => _divideToolTip;
            set
            {
                _divideToolTip = value;
                NotifyPropertyChanged();
            }
        }

        public string MergeToolTip => $"Merge {BexConstants.UmbrellaTypeName.ToLower()}s";
        
        public bool IsMergeButtonEnabled
        {
            get => _isMergeButtonEnabled;
            set
            {
                _isMergeButtonEnabled = value;
                NotifyPropertyChanged();
            }
        }
        public bool IsMergeButtonVisible
        {
            get => _isMergeButtonVisible;
            set
            {
                _isMergeButtonVisible = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsMoveLeftButtonVisible
        {
            get => _isMoveLeftButtonVisible;
            set
            {
                _isMoveLeftButtonVisible = value; 
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
    }

    
}


