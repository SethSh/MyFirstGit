using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MunichRe.Bex.ApiClient.ClientApi;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using SubmissionCollector.Extensions;
using SubmissionCollector.View.Enums;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for UmbrellaTypeWizard.xaml
    /// </summary>
    public partial class UmbrellaTypePolicyProfileWizard
    {
        public FormResponse Response { get; set; }
        private readonly UmbrellaTypePolicyProfileWizardViewModel _viewModel;
        private readonly GridLength _zeroStar = new GridLength(0, GridUnitType.Star);
        private readonly GridLength _oneStar = new GridLength(1, GridUnitType.Star);
        private const int BorderThicknessWithEmphasis = 5;

        public UmbrellaTypePolicyProfileWizard(UmbrellaTypePolicyProfileWizardViewModel viewModel)
        {
            _viewModel = viewModel;
            InitializeComponent();

            DataContext = _viewModel;

            RedrawForm();
            RedrawButtons();
            Response = FormResponse.Cancel;
        }
        
        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            Response = FormResponse.Ok;
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            Response = FormResponse.Cancel;
        }
        
        private void DivideButton_OnClick(object sender, RoutedEventArgs e)
        {
            var templateView = (UmbrellaComponentView)((Button)sender).DataContext;
            var visibleCountAtStart = _viewModel.ComponentViews.Count(x => x.IsVisible);

            //step 1 - identify new views 
            var newViews = new List<UmbrellaComponentView>();
            for (var i = 0; i < templateView.UmbrellaTypes.Count - 1; i++)
            {
                newViews.Add(_viewModel.ComponentViews.Single(x => x.DisplayOrder == visibleCountAtStart + i));
            }

            //step 2 - move other visible views to right
            foreach (var item in _viewModel.ComponentViews.Where(x => x.IsVisible && x.DisplayOrder > templateView.DisplayOrder))
            {
                item.DisplayOrder += templateView.UmbrellaTypes.Count-1;
            }

            //step 3 - change new views
            var counter = 0;
            var otherUmbrellaTypes = templateView.UmbrellaTypes.Where(x => x.UmbrellaTypeName != templateView.UmbrellaTypes.First().UmbrellaTypeName).ToList();
            foreach (var view in newViews)
            {
                view.DisplayOrder = templateView.DisplayOrder + 1 + counter;
                var umbrellaType = otherUmbrellaTypes[counter];
             
                view.Sublines = templateView.Sublines;
                view.Id = ++_viewModel.PolicyProfilesMaxIndex;
                view.UmbrellaTypes = new ObservableCollection<UmbrellaTypeViewModel> { umbrellaType };
                view.IsVisible = true;
                view.IsNew = true;
                view.IsLocked = false;
                view.Guid = Guid.NewGuid();
                view.Name = BexConstants.PolicyProfileName;
                view.FullName = $"{BexConstants.PolicyProfileName} " +
                                $"{UmbrellaTypesFromBex.AbbreviateUmbrellaType(umbrellaType.UmbrellaTypeName)}";

                counter++;
            }
            
            templateView.UmbrellaTypes.RemoveAllButFirst();
            templateView.FullName = $"{BexConstants.PolicyProfileName} " +
                                    $"{UmbrellaTypesFromBex.AbbreviateUmbrellaType(templateView.UmbrellaTypes.First().UmbrellaTypeName)}";


            RedrawForm();
            RedrawButtons();

            var matches = FindSublineMatches(templateView, true);
            HighlightComponentsForAWhile(matches.ToList(), 2);
        }

        private void MergeButton_OnClick(object sender, RoutedEventArgs e)
        {
            var templateView = (UmbrellaComponentView)((Button)sender).DataContext;

            var siblingViews = _viewModel.ComponentViews.Where(x => x.Guid != templateView.Guid && x.IsVisible &&
                                                                    x.Sublines.IsEqualTo(templateView.Sublines)).ToList();

            templateView.UmbrellaTypes = new ObservableCollection<UmbrellaTypeViewModel>(_viewModel.SegmentUmbrellaTypes);
            templateView.FullName = BexConstants.PolicyProfileName;

            var counter = 0;
            foreach (var view in _viewModel.ComponentViews.Where(vm => vm.IsVisible).Except(siblingViews).OrderBy(vm => vm.DisplayOrder))
            {
                view.DisplayOrder = counter++;
            }
            
            foreach (var view in siblingViews.OrderBy(x => x.DisplayOrder))
            {
                view.DisplayOrder = counter++;
                view.IsVisible = false;
                view.Sublines = null;
                view.UmbrellaTypes = null;
                view.Name = BexConstants.PolicyProfileName;
            }

            RedrawForm();
            RedrawButtons();
        }

        private IEnumerable<UmbrellaComponentView> FindSublineMatches(UmbrellaComponentView componentView, bool includeSelf = false)
        {
            if (includeSelf)
            {
                return _viewModel.ComponentViews.Where(x => x.IsVisible && x.Sublines.IsEqualTo(componentView.Sublines));
            }

            return _viewModel.ComponentViews.Where(x => x.IsVisible
                                                                  && x.Id != componentView.Id
                                                                  && x.Sublines.IsEqualTo(componentView.Sublines));
        }

        private void RefreshDivideButtons()
        {
            foreach (var item in _viewModel.ComponentViews.Where(x => x.IsVisible))
            {
                var divideButtonMessage = new StringBuilder();
                
                if (item.Sublines.Any(x => x.IsPersonal))
                {
                    item.IsDivideButtonEnabled = false;
                    item.IsMergeButtonEnabled = false;
                    divideButtonMessage.AppendLine($"Not applicable when a personal {BexConstants.SublineName.ToLower()} is present");
                }

                var profileCount = _viewModel.ComponentViews.Count(x => x.IsVisible);
                var notEnoughPlaceholderProfiles = profileCount + (item.UmbrellaTypes.Count - 1) > BexConstants.MaximumNumberOfDataComponents;
                if (notEnoughPlaceholderProfiles)
                {
                    divideButtonMessage.AppendLine($"Not enough {BexConstants.PolicyProfileName.ToLower()}s available to handle a divide");
                }
                
                var anyMatches = FindSublineMatches(item).Any();
                if (!anyMatches && item.UmbrellaTypes.Count == 1)
                {
                    item.IsDivideButtonEnabled = false;
                    item.IsMergeButtonEnabled = false;
                    divideButtonMessage.AppendLine($"Not applicable when only one {BexConstants.UmbrellaTypeName.ToLower()}");
                }

                item.IsDivideButtonEnabled = !anyMatches && divideButtonMessage.Length == 0;
                item.IsMergeButtonEnabled = anyMatches && divideButtonMessage.Length == 0;

                item.DivideToolTip = divideButtonMessage.Length > 0 
                    ? $"Divide {BexConstants.UmbrellaTypeName.ToLower()}s\n" + divideButtonMessage 
                    : $"Divide {BexConstants.UmbrellaTypeName.ToLower()}s";
                
            }
        }

        private void ShowMergeButtons()
        {
            foreach (var item in _viewModel.ComponentViews.Where(x => x.IsVisible))
            {
                if (item.Sublines.Any(subline => subline.IsPersonal))
                {
                    item.IsMergeButtonVisible = false;
                    item.IsDivideButtonVisible = true;
                    continue;
                }

                var isMerged = item.UmbrellaTypes.Count == _viewModel.SegmentUmbrellaTypes.Count;
                item.IsMergeButtonVisible = !isMerged;
                item.IsDivideButtonVisible = isMerged;
            }
        }

        private void RedrawButtons()
        {
            ShowMoveButtons();
            RefreshDivideButtons();
            ShowMergeButtons();
        }

        private void ShowMoveButtons()
        {
            foreach (var item in _viewModel.ComponentViews.Where(x => x.IsVisible))
            {
                item.IsMoveLeftButtonVisible = _viewModel.GetIsMoveLeftButtonVisible(item.DisplayOrder);
                item.IsMoveRightButtonVisible = _viewModel.GetIsMoveRightButtonVisible(item.DisplayOrder);
            }
        }

        private void RedrawForm()
        {
            var count = _viewModel.ComponentViews.Count(x => x.IsVisible);
            if (count == 0) count = 1;

            OverallLabel.SetValue(Grid.ColumnSpanProperty, count);

            PolicyProfile2Column.Width = GetStar(count, 2);
            PolicyProfile3Column.Width = GetStar(count, 3);
            PolicyProfile4Column.Width = GetStar(count, 4);
            PolicyProfile5Column.Width = GetStar(count, 5);
            PolicyProfile6Column.Width = GetStar(count, 6);
        }

        private static void HighlightComponentsForAWhile(IList<UmbrellaComponentView> components, int seconds)
        {
            components.ForEach(x => x.BorderThickness = BorderThicknessWithEmphasis);
            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();

            var task = new Task(() => WaitSeconds(seconds));
            task.ContinueWith(task1 =>
            {
                components.ForEach(x => x.BorderThickness = BaseUmbrellaPolicyProfileWizardViewModel.DefaultBorderThickness);
            }, CancellationToken.None, TaskContinuationOptions.None, scheduler);

            task.Start();
        }

        private static void WaitSeconds(int secondCount)
        {
            Thread.Sleep(secondCount * 1000);
        }
        
        private void MergeButton_MouseEnter(object sender, MouseEventArgs e)
        {
            var templateView = (UmbrellaComponentView)((Button)sender).DataContext;
            FindSublineMatches(templateView, true).ForEach(x => x.BorderThickness = BorderThicknessWithEmphasis);
        }
        
        private void MergeButton_MouseLeave(object sender, MouseEventArgs e)
        {
            var templateView = (UmbrellaComponentView)((Button)sender).DataContext;
            FindSublineMatches(templateView, true).ForEach(x => x.BorderThickness = BaseUmbrellaPolicyProfileWizardViewModel.DefaultBorderThickness);
        }

        private GridLength GetStar(int index, int policyProfileId)
        {
            return index >= policyProfileId ? _oneStar : _zeroStar;
        }


        private void MoveLeft(object sender, RoutedEventArgs e)
        {
            var componentView = (UmbrellaComponentView)((Button)sender).DataContext;
            var displayOrder = componentView.DisplayOrder;
            
            var componentViewOneToLeft = _viewModel.ComponentViews.Single(x => x.DisplayOrder == displayOrder - 1);
            SwitchDisplayOrders(componentView, componentViewOneToLeft);

            RedrawButtons();
            RedrawForm();
        }

        private void MoveRight(object sender, RoutedEventArgs e)
        {
            var componentView = (UmbrellaComponentView)((Button)sender).DataContext;
            var displayOrder = componentView.DisplayOrder;
            
            var componentViewOneToRight = _viewModel.ComponentViews.Single(x => x.DisplayOrder == displayOrder + 1);
            SwitchDisplayOrders(componentView, componentViewOneToRight);

            RedrawButtons();
            RedrawForm();
        }

        private static void SwitchDisplayOrders(UmbrellaComponentView view1, UmbrellaComponentView view2)
        {
            var displayOrder1 = view1.DisplayOrder;
            var displayOrder2 = view2.DisplayOrder;

            view2.DisplayOrder = displayOrder1;
            view1.DisplayOrder = displayOrder2;
        }
    }
}
