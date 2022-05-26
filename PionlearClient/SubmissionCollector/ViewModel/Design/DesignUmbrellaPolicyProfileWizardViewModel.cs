using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Media.Imaging;
using MunichRe.Bex.ApiClient.ClientApi;
using PionlearClient;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.Properties;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignUmbrellaPolicyProfileWizardViewModel : BaseUmbrellaPolicyProfileWizardViewModel
    {
        public DesignUmbrellaPolicyProfileWizardViewModel()
        {
            Title = $"Submission Segment Display Name {BexConstants.Dash} {BexConstants.PolicyProfileName}s";
            
            SegmentUmbrellaTypes = new List<UmbrellaTypeViewModel>
            {
                new UmbrellaTypeViewModel {UmbrellaTypeCode = 1, UmbrellaTypeName = "Supported"},
                new UmbrellaTypeViewModel {UmbrellaTypeCode = 2, UmbrellaTypeName = "Unsupported"},
                new UmbrellaTypeViewModel {UmbrellaTypeCode = 3, UmbrellaTypeName = "Unsupported Excess"}
            };

            var s1 = new FakeSubline {ImageSource = Resources.AngryBird.ToBitmapSource()};
            var s2 = new FakeSubline {ImageSource = Resources.AngryBird.ToBitmapSource()};
            var sublines = new ObservableCollection<ISubline>(new List<ISubline> {s1, s2});

            ComponentViews = new List<UmbrellaComponentView>();
            for (var i = 0; i < BexConstants.MaximumNumberOfDataComponents; i++)
            {
                var umbrellaWizardView = new UmbrellaComponentView
                {
                    Id = i,
                    DisplayOrder = i,
                    Name = $"{BexConstants.PolicyProfileName}",
                    IsVisible = true,
                    IsLocked = true,
                    IsShaded= i == 0,
                    IsNew = true,
                    HasData = true,
                    BorderThickness = 1,
                    IsMoveLeftButtonVisible = true,
                    IsMoveRightButtonVisible = true,
                    IsMergeButtonEnabled = true,
                    IsMergeButtonVisible = true,
                    IsDivideButtonEnabled = true,
                    IsDivideButtonVisible = true,
                    Sublines = new ObservableCollection<ISubline>(sublines),
                    UmbrellaTypes = new ObservableCollection<UmbrellaTypeViewModel>
                    {
                        new UmbrellaTypeViewModel {UmbrellaTypeCode = 1, UmbrellaTypeName = "Supported"},
                        new UmbrellaTypeViewModel {UmbrellaTypeCode = 2, UmbrellaTypeName = "Unsupported"},
                        new UmbrellaTypeViewModel {UmbrellaTypeCode = 3, UmbrellaTypeName = "Unsupported Excess"}
                    }
                };
                
                ComponentViews.Add(umbrellaWizardView);
            }

        }

    }

    internal class FakeSubline : ISubline
    {
        public FakeSubline()
        {
            IsWorkersComp = false;
            HasPolicyProfile = true;
        }
        public int DisplayOrder { get; set; }
        public bool IsSelected { get; set; }
        public bool IsExpanded { get; set; }
        public string Name => "Subline of Business";
        public int SegmentId { get; set; }
        public string ShortName => "Sublob";
        public string ShortNameWithLob => $"LOB {BexConstants.Dash} Sublob";
        public string NameWithLob => $"Line of Business {BexConstants.Dash} Subline of Business";
        public string LobShortName => "LOB";
        public bool HasPolicyProfile { get; }
        public bool HasStateProfile { get; set; }
        public bool HasHazardProfile { get; set; }
        public bool IsLineExclusive { get; set; }
        public int Code => 0;
        public BitmapSource ImageSource { get; set; }
        public bool IsPersonal => true;
        public ISegment FindParentSegment()
        {
            throw new NotImplementedException();
        }

        public bool IsWorkersComp { get; }

        public LineOfBusinessType LineOfBusinessType { get; set; }
        
    }
}
