using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Comparers;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;
using Button = System.Windows.Controls.Button;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using ListBox = System.Windows.Controls.ListBox;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for SublineWizard.xaml
    /// </summary>
    public partial class SublineWizard
    {
        public FormResponse Response { get; set; }
        private readonly SublineWizardViewModel _viewModel;
        private ListBox _dragSource;
        private Type _myDataType;
        private readonly GridLength _zeroPixel = new GridLength(0, GridUnitType.Pixel);
        private readonly GridLength _oneStar = new GridLength(1, GridUnitType.Star);
        private const int BorderThicknessWithEmphasis = 3;
        private readonly long _personalUmbrellaTypeCode;
        private readonly string _workersCompensationName;
        private readonly string _autoPhysicalDamageName;
        private readonly string _propertyName;

        public SublineWizard(SublineWizardViewModel viewModel)
        {
            InitializeComponent();

            _workersCompensationName = SublineCodesFromBex.GetLobName(BexConstants.WorkersCompensationIndemnitySublineCode);
            _autoPhysicalDamageName = SublineCodesFromBex.GetLobName(BexConstants.AutoPhysicalDamageCommercialSublineCode);
            _propertyName = SublineCodesFromBex.GetLobName(BexConstants.CommercialPropertySublineCode);
            _personalUmbrellaTypeCode = UmbrellaTypesFromBex.GetPersonalCode();

            _viewModel = viewModel;

            Response = FormResponse.Cancel;
            DataContext = _viewModel;
            Available.Focus();

            Loaded += SublineSelectorWizard_Loaded;
        }

        private void SublineSelectorWizard_Loaded(object sender, RoutedEventArgs e)
        {
            RedrawForm();
        }

        
        private void Available_OnDrop(object sender, DragEventArgs e)
        {
            if (_dragSource.Name == ((ListBox) sender).Name) return;

            var subline = (ISubline) e.Data.GetData(_myDataType);
            Debug.Assert(subline != null, $"{BexConstants.SublineName} can't be found");
            
            _viewModel.SegmentSublines.RemoveSubline(subline);

            _viewModel.AvailableSublines.Add(subline);
            _viewModel.AvailableSublines.Sort(x => x.ShortNameWithLob);

            foreach (var policy in _viewModel.PolicyProfiles)
            {
                policy.Sublines.RemoveSubline(subline);
            }

            if (!_viewModel.SegmentSublines.Any(x => x.HasPolicyProfile))
            {
                _viewModel.PolicyProfiles.ForEach(x => x.IsVisible = false);
            }


            foreach (var state in _viewModel.StateProfiles)
            {
                state.Sublines.RemoveSubline(subline);
            }

            if (!_viewModel.SegmentSublines.ContainsState())
            {
                _viewModel.StateProfiles.ForEach(x => x.IsVisible=false);
            }

            
            
            _viewModel.ExposureSets.ForEach(exposure => exposure.Sublines.RemoveSubline(subline));
            _viewModel.AggregateLossSets.ForEach(aggregateLoss => aggregateLoss.Sublines.RemoveSubline(subline));
            _viewModel.IndividualLossSets.ForEach(individualLoss => individualLoss.Sublines.RemoveSubline(subline));
            _viewModel.RateChangeSets.ForEach(rateChange => rateChange.Sublines.RemoveSubline(subline));

            if (!_viewModel.SegmentSublines.Any())
            {
                _viewModel.ComponentViews.ForEach(x => x.IsVisible = false);
                _viewModel.InitializeMaxIndices();
            } 
            RedrawForm();
        }

        
        private bool IsSublineCombinationOk(ISubline subline)
        {
            if (!_viewModel.SegmentSublines.Any()) return true;
            if (!subline.IsLineExclusive) return _viewModel.SegmentSublines.All(s => !s.IsLineExclusive);
            
            var allLinesExclusive = _viewModel.SegmentSublines.All(s => s.IsLineExclusive);

            var possibleNewLinesOfBusiness = _viewModel.SegmentSublines.Select(s => s.LineOfBusinessType).ToList();
            possibleNewLinesOfBusiness.Add(subline.LineOfBusinessType);
            var sameLineOfBusiness = possibleNewLinesOfBusiness.Distinct().Count() == 1;
            
            return allLinesExclusive && sameLineOfBusiness;
        }

        private void Available_OnDrag(object sender, DragEventArgs e)
        {
            if (_dragSource.Name == "Segment") return;
            
            if (_dragSource.Name == "Available")
            {
                var subline = (ISubline) e.Data.GetData(_myDataType);
                Debug.Assert(subline != null, $"{BexConstants.SublineName} can't be found");

                if (!IsSublineCombinationOk(subline))
                {
                    _viewModel.Log = $"Can't add the {subline.ShortNameWithLob} {BexConstants.SublineName.ToLower()} " +
                                     $"to the {BexConstants.SegmentName} {BexConstants.SublineName}s.";
                    e.Effects = DragDropEffects.None;
                    e.Handled = true;
                }

                if (_viewModel.Segment.IsUmbrella && !IsUmbrellaOkay(subline.LineOfBusinessType))
                {
                    var lineOfBusinessName = MapToLineOfBusinessName(subline.LineOfBusinessType);
                    _viewModel.Log = $"Can't include a {lineOfBusinessName} {BexConstants.SublineName.ToLower()} " +
                                     $"in an {BexConstants.UmbrellaName.ToLower()} {BexConstants.SegmentName.ToLower()}.";
                    e.Effects = DragDropEffects.None;
                    e.Handled = true;
                }
                return;
            }
        
            var componentView = (ComponentView)_dragSource.DataContext;
            _viewModel.Log = $"Can't drag and drop a {_viewModel.ComponentViewNamePrefixes[componentView.Type].ToLower()} " +
                             $"{BexConstants.SublineName.ToLower()} into available {BexConstants.SublineName.ToLower()}s";
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            _viewModel.Log = string.Empty;
        }
        
        private void Segment_OnDrop(object sender, DragEventArgs e)
        {
            if (_dragSource.Name == ((ListBox)sender).Name) return;

            var subline = (ISubline)e.Data.GetData(_myDataType);
            Debug.Assert(subline != null, $"{BexConstants.SublineName} can't be found");

            if (_dragSource.Name == "Available")
            {
                if (!IsSublineCombinationOk(subline))
                {
                    _viewModel.Log = $"Can't add the {subline.ShortNameWithLob} {BexConstants.SublineName.ToLower()} " +
                                     $"to the {BexConstants.SegmentName} {BexConstants.SublineName}s.";
                    e.Effects = DragDropEffects.None;
                    e.Handled = true;
                    return;
                }

                
                if (_viewModel.Segment.IsUmbrella && !IsUmbrellaOkay(subline.LineOfBusinessType))
                {
                    var lineOfBusinessName = MapToLineOfBusinessName(subline.LineOfBusinessType);
                    _viewModel.Log = $"Can't include a {lineOfBusinessName} {BexConstants.SublineName.ToLower()} " +
                                     $"in an {BexConstants.UmbrellaName.ToLower()} {BexConstants.SegmentName.ToLower()}.";
                    e.Effects = DragDropEffects.None;
                    e.Handled = true;
                    return;
                }
            }

            SendSublineFromAvailable(subline);
            _viewModel.Log = string.Empty;
        }

        private void Segment_OnDrag(object sender, DragEventArgs e)
        {
            if (_dragSource.Name == "Available" || _dragSource.Name == "Segment") return;

            var componentView = (ComponentView)_dragSource.DataContext;
            _viewModel.Log = $"Can't drag and drop a {_viewModel.ComponentViewNamePrefixes[componentView.Type].ToLower()} " +
                             $"{BexConstants.SublineName.ToLower()} into {BexConstants.SegmentName.ToLower()} " +
                             $"{BexConstants.SublineName.ToLower()}s";
            e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void ComponentViewListBox_OnDrag(object sender, DragEventArgs e)
        {
            var componentView = (ComponentView)((ListBox)sender).DataContext;
            var type = componentView.Type;

            if (componentView.UmbrellaType.HasValue && componentView.UmbrellaType.Value != _personalUmbrellaTypeCode)
            {
                _viewModel.Log = $"Can't drag and drop from or into a {BexConstants.PolicyProfileName.ToLower()} " +
                                 $"that's been separated by {BexConstants.UmbrellaTypeName.ToLower()}.";
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            if (_dragSource.Name == "Available")
            {
                _viewModel.Log = $"Can't drag and drop an available {BexConstants.SublineName.ToLower()} " +
                                 $"into {_viewModel.ComponentViewNamePrefixes[componentView.Type].ToLower()}";
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            if (_dragSource.Name == "Segment")
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }

            var otherComponentView = (ComponentView)_dragSource.DataContext;
            var otherType = otherComponentView.Type;

            if (otherComponentView.UmbrellaType.HasValue && otherComponentView.UmbrellaType.Value != _personalUmbrellaTypeCode)
            {
                _viewModel.Log = $"Can't drag and drop from or into a {BexConstants.PolicyProfileName.ToLower()} " +
                                 $"that's been separated by {BexConstants.UmbrellaTypeName.ToLower()}.";
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            if (type == otherType)
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }

            _viewModel.Log = $"Can't drag and drop a {_viewModel.ComponentViewNamePrefixes[otherType].ToLower()} " +
                             $"{BexConstants.SublineName.ToLower()} into {_viewModel.ComponentViewNamePrefixes[type].ToLower()}";
            e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void ComponentViewListBox_OnDrop(object sender, DragEventArgs e)
        {
            var componentView = (ComponentView) ((ListBox) sender).DataContext;
            var type = componentView.Type;

            var subline = (ISubline) e.Data.GetData(_myDataType);
            Debug.Assert(subline != null, "Nothing to drop");

            var isDragFromSegment = _dragSource.Name == "Segment";
            var alreadyExists = isDragFromSegment && GetComponentViewsSublines(type).ContainsSubline(subline);
            if (alreadyExists) return;

            if (isDragFromSegment)
            {
                if (componentView.Type == ComponentViewType.Policy && !subline.HasPolicyProfile)
                {
                    _viewModel.Log = $"{subline.ShortNameWithLob} doesn't allow a {BexConstants.PolicyProfileName.ToLower()}.";
                    return;
                }

                if (componentView.Type == ComponentViewType.State && !subline.HasStateProfile)
                {
                    _viewModel.Log = $"{subline.ShortNameWithLob} doesn't allow a {BexConstants.StateProfileName.ToLower()}.";
                    return;
                }
            }

            var existingSublines = _viewModel.ComponentViews.First(x => x.Type == type && x.DisplayOrder == componentView.DisplayOrder).Sublines;

            if (!isDragFromSegment)
            {
                var draggedComponentView = (ComponentView) _dragSource.DataContext;
                var draggedType = draggedComponentView.Type;

                alreadyExists = type == draggedType && existingSublines.ContainsSubline(subline);
                if (alreadyExists) return;

                ((ObservableCollection<ISubline>)_dragSource.ItemsSource).RemoveSubline(subline);
            }

            existingSublines.Add(subline);
            existingSublines.Sort(x => x.ShortNameWithLob);
            _viewModel.Log = string.Empty;

            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void TrashImage_OnDrag(DragEventArgs e)
        {
            var isOk = true;
            if (_dragSource.Name == "Available")
            {
                _viewModel.Log = $"Can't drag and drop an available {BexConstants.SublineName.ToLower()} into the trash.";
                isOk = false;
            }

            if (_dragSource.Name == "Segment")
            {
                _viewModel.Log = $"Can't drag and drop a {BexConstants.SegmentName.ToLower()} {BexConstants.SublineName.ToLower()} into the trash.";
                isOk = false;
            }

            if (_dragSource.DataContext is ComponentView componentView && componentView.Type == ComponentViewType.Policy && componentView.UmbrellaType != null)
            {
                _viewModel.Log = $"Can't drag and drop a {BexConstants.SublineName.ToLower()} " +
                                 $"from a {BexConstants.PolicyProfileName.ToLower()} that's been separated by {BexConstants.UmbrellaTypeName.ToLower()} into the trash.";
                isOk = false;
            }

            if (isOk) return;
            e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void TrashImage_OnDragOver(object sender, DragEventArgs e)
        {
            TrashImage_OnDrag(e);
        }

        private void TrashImage_OnDragEnter(object sender, DragEventArgs e)
        {
            TrashImage_OnDrag(e);
        }

        private void TrashImage_OnDrop(object sender, DragEventArgs e)
        {
            var data = (ISubline)e.Data.GetData(_myDataType);

            ((ObservableCollection<ISubline>)_dragSource.ItemsSource).RemoveSubline(data);
            _viewModel.Log = string.Empty;

            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void AddComponentViewClick(object sender, RoutedEventArgs e)
        {
            var componentView = (ComponentView)((Button)sender).DataContext;
            var displayOrder = componentView.DisplayOrder;
            var type = componentView.Type;

            var view = _viewModel.ComponentViews.Where(x => x.Type == type);
            _viewModel.MaximumIndices[type] = _viewModel.MaximumIndices[type].AddOne();
            
            // ReSharper disable once PossibleInvalidOperationException
            AddComponent(view, displayOrder, _viewModel.MaximumIndices[type].Value, _viewModel.ComponentViewNamePrefixes[type]);
        }

        private void AddComponent(IEnumerable<ComponentView> componentViews, int index, string namePrefix)
        {
            var component = componentViews.First(x => x.DisplayOrder == 0);
            SetNewComponentViewProperties(component, index, 0, namePrefix);
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void AddComponent(IEnumerable<ComponentView> componentViews, int displayOrder, int index, string namePrefix)
        {
            var list = componentViews.ToList();
            var newComponent = list.Single(x => x.DisplayOrder == BexConstants.MaximumNumberOfDataComponents-1);
            foreach (var item in list.Where(x => x.DisplayOrder > displayOrder && x.DisplayOrder < _viewModel.MaxDisplayOrder))
            {
                item.DisplayOrder++;
            }
            SetNewComponentViewProperties(newComponent, index, displayOrder + 1, namePrefix);
            HighlightComponentsForAWhile(newComponent, 2);
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private static void HighlightComponentsForAWhile(ComponentView componentView, int seconds)
        {
            componentView.BorderThickness = BorderThicknessWithEmphasis;

            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            var task = new Task(() => WaitSeconds(seconds));
            task.ContinueWith(task1 =>
            {
                componentView.BorderThickness = BaseSublineSelectorWizardViewModel.DefaultBorderThickness;
            }, CancellationToken.None, TaskContinuationOptions.None, scheduler);

            task.Start();
        }

        private static void WaitSeconds(int secondCount)
        {
            Thread.Sleep(secondCount * 1000);
        }

        private void SetNewComponentViewProperties(ComponentView componentView, int index, int displayOrder, string namePrefix)
        {
            componentView.Id = index;
            componentView.Name = $"{namePrefix}";
            componentView.IsVisible = true;
            componentView.DisplayOrder = displayOrder;
            componentView.Sublines = new ObservableCollection<ISubline>();
            componentView.IsMoveRightButtonVisible = _viewModel.GetIsMoveRightButtonVisible(componentView.Type, displayOrder);
            componentView.IsMoveLeftButtonVisible = _viewModel.GetIsMoveLeftButtonVisible(displayOrder);
            componentView.Guid = Guid.NewGuid();
            componentView.IsNew = true;
            componentView.HasData = false;
            componentView.IsLocked = false;
            
            RedrawForm();
        }

        private void DeleteComponentViewClick(object sender, RoutedEventArgs e)
        {
            var componentView = (ComponentView) ((Button) sender).DataContext;
            var displayOrder = componentView.DisplayOrder;
            var type = componentView.Type;

            if (componentView.Sublines.Any())
            {
                MessageHelper.Show($"Remove {BexConstants.SublineName.ToLower()}s before deleting item.", MessageType.Stop);
                return;
            }
            
            DeleteComponent(_viewModel.ComponentViews.Where(x => x.Type == type), displayOrder);
            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void DeleteComponent(IEnumerable<ComponentView> componentView, int displayOrder)
        {
            var list = componentView.ToList();
            
            var thisComponentView = list.Single(x => x.DisplayOrder == displayOrder);
            thisComponentView.IsVisible = false;
            thisComponentView.Sublines = null;
            thisComponentView.DisplayOrder = _viewModel.MaxDisplayOrder;
            if (thisComponentView.IsNew) _viewModel.MaximumIndices[thisComponentView.Type]--;

            foreach (var item in list.Where(x => x.Guid != thisComponentView.Guid && x.DisplayOrder > displayOrder))
            {
                item.DisplayOrder--;
            }
            
            RedrawForm();
        }
        
        private IEnumerable<ISubline> GetComponentViewsSublines(ComponentViewType type)
        {
            return _viewModel.ComponentViews.Where(x => x.Type == type &&  x.Sublines != null).SelectMany(x => x.Sublines);
        }
        
        private static object GetDataFromListBox(ListBox source, Point point)
        {
            var element = source.InputHitTest(point) as UIElement;
            if (element == null) return null;

            var data = DependencyProperty.UnsetValue;
            while (data == DependencyProperty.UnsetValue)
            {
                if (element == null) continue;

                data = source.ItemContainerGenerator.ItemFromContainer(element);
                if (data == DependencyProperty.UnsetValue)
                {
                    element = VisualTreeHelper.GetParent(element) as UIElement;
                }
                if (Equals(element, source))
                {
                    return null;
                }
            }
            return data != DependencyProperty.UnsetValue ? data : null;
        }
        
        private void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var parent = (ListBox) sender;
            _dragSource = parent;
            var data = GetDataFromListBox(_dragSource, e.GetPosition(parent));

            if (data == null) return;

            _myDataType = data.GetType();
            DragDrop.DoDragDrop(parent, data, DragDropEffects.All);
        }
        
        private void RedrawForm()
        {
            _viewModel.SetSublineCustomization();
            _viewModel.SelectedSublinesRowSpan = _viewModel.GetSelectedSublineRowSpan();

            if (_viewModel.AreSublinesRigid)
            {
                _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
                return;
            }

            var hasProperty = _viewModel.SegmentSublines.ContainsProperty();
            var hasPolicy = _viewModel.SegmentSublines.ContainsPolicy();
            var hasState = _viewModel.SegmentSublines.ContainsState();
            var hasCommercial = _viewModel.SegmentSublines.ContainsCommercial();
            
            _viewModel.PolicyRowHeight = !hasProperty && hasPolicy ? _oneStar : _zeroPixel;
            _viewModel.StateRowHeight = !hasProperty && hasState ? _oneStar : _zeroPixel;
            
            _viewModel.OccupancyTypeRowHeight = hasProperty && hasCommercial ? _oneStar : _zeroPixel; 
            _viewModel.TotalInsuredValueRowHeight = hasProperty ? _oneStar : _zeroPixel;

            _viewModel.ExposureRowHeight = _oneStar;
            _viewModel.AggregateLossRowHeight = _oneStar;
            _viewModel.IndividualLossRowHeight = _oneStar;
            _viewModel.RateChangeRowHeight = _oneStar;

            var countByComponentType = _viewModel.GetComponentCountByType();
            var maximumCount = countByComponentType.Values.Max();
            if (maximumCount == 0) maximumCount = 1;

            DataComponent1Column.Width = GetGridLength(maximumCount, 1);
            DataComponent2Column.Width = GetGridLength(maximumCount, 2);
            DataComponent3Column.Width = GetGridLength(maximumCount, 3);
            DataComponent4Column.Width = GetGridLength(maximumCount, 4);
            DataComponent5Column.Width = GetGridLength(maximumCount, 5);
            DataComponent6Column.Width = GetGridLength(maximumCount, 6);
            
            foreach (var componentView in _viewModel.ComponentViews.Where(x => x.IsVisible))
            {
                componentView.IsMoveRightButtonVisible = _viewModel.GetIsMoveRightButtonVisible(componentView.Type, componentView.DisplayOrder);
                componentView.IsMoveLeftButtonVisible = _viewModel.GetIsMoveLeftButtonVisible(componentView.DisplayOrder);
                componentView.IsAddButtonVisible = countByComponentType[componentView.Type] < BexConstants.MaximumNumberOfDataComponents;
                componentView.IsDeleteButtonVisible = countByComponentType[componentView.Type] > 1 && !(componentView.UmbrellaType.HasValue  && componentView.UmbrellaType.Value != _personalUmbrellaTypeCode);
            }

            _viewModel.ComponentViews.Where(x => x.IsVisible).ForEach(x => x.BorderThickness = BaseSublineSelectorWizardViewModel.DefaultBorderThickness);
            _viewModel.ComponentViews.Where(x => !x.IsVisible).ForEach(x => x.BorderThickness = 0);

            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }
        
        private GridLength GetGridLength(int displayOrder, int dataComponentIndex)
        {
            return displayOrder >= dataComponentIndex ? _oneStar : _zeroPixel;
        }
        
        private void ClearLog(object sender, DragEventArgs e)
        {
            _viewModel.Log = string.Empty;
        }
        
        
        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            Response = FormResponse.Ok;
        }

        private bool GetOkButtonEnabledStatus()
        {
            _viewModel.OkButtonToolTip = null;

            var sb = new StringBuilder();

            if (!_viewModel.SegmentSublines.Any())
            {
                sb.AppendLine($"There must be at least one {BexConstants.SegmentName.ToLower()} {BexConstants.SublineName.ToLower()}");
            }

            CheckAtLeastOneSubline(sb);
            CheckLineExclusiveIsOk(sb);
            
            if (sb.Length <= 0) return true;

            _viewModel.OkButtonToolTip = sb.ToString();
            return false;
        }

      
        private void CheckAtLeastOneSubline(StringBuilder sb)
        {
            _viewModel.ComponentViews.Where(view => view.IsVisible && !view.Sublines.Any()).ForEach(view =>
            {
                var componentName = _viewModel.ComponentViewNamePrefixes[view.Type];
                var message = $"Add at least one {BexConstants.SublineName.ToLower()} to (or delete) the empty {componentName.ToLower()}";
                sb.AppendLine(message);
            });
        }

        private void CheckLineExclusiveIsOk(StringBuilder sb)
        {
            var isExclusiveSubline = _viewModel.SegmentSublines.Any(s => s.IsLineExclusive);
            var isNonExclusiveSubline = _viewModel.SegmentSublines.Any(s => !s.IsLineExclusive);
            if (isExclusiveSubline && isNonExclusiveSubline)
            {
                var message = $"The {BexConstants.SegmentName.ToLower()} contains {BexConstants.SublineName.ToLower()}s " +
                              "that aren't combine-able";
                sb.AppendLine(message);
            }
        }

        
        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            Response = FormResponse.Cancel;
        }
        
        private static bool ContainsSubline(IEnumerable<ComponentView> views, ISubline subline)
        {
            return views.Any(item => item.Sublines?.ContainsSubline(subline) == true);
        }

        

        private void Available_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (Available.SelectedValue == null) return;
            var subline = (ISubline) Available.SelectedValue;

            if (!IsSublineCombinationOk(subline))
            {
                _viewModel.Log = $"Can't add the {subline.ShortNameWithLob} {BexConstants.SublineName.ToLower()} " +
                                 $"to the {BexConstants.SegmentName} {BexConstants.SublineName}s.";
                return;
            }

            if (_viewModel.Segment.IsUmbrella && !IsUmbrellaOkay(subline.LineOfBusinessType))
            {
                var lineOfBusinessName = MapToLineOfBusinessName(subline.LineOfBusinessType);
                _viewModel.Log = $"Can't include a {lineOfBusinessName} {BexConstants.SublineName.ToLower()} " +
                                 $"in an {BexConstants.UmbrellaName.ToLower()} {BexConstants.SegmentName.ToLower()}s.";
                return;
            }

            SendSublineFromAvailable(subline);
            _viewModel.Log = string.Empty;
        }

        private void SendSublineFromAvailable(ISubline subline)
        {
            _viewModel.SegmentSublines.Add(subline);
            _viewModel.AvailableSublines.RemoveSubline(subline);
            _viewModel.SegmentSublines.Sort(x => x.ShortNameWithLob);

            if (!_viewModel.AutomaticallySyncSublines) return;

            ComponentViewType type;
            var names = _viewModel.ComponentViewNamePrefixes;
            if (subline.HasPolicyProfile && !ContainsSubline(_viewModel.PolicyProfiles, subline))
            {
                type = ComponentViewType.Policy;
                if (!_viewModel.PolicyProfiles.Any(x => x.IsVisible))
                {
                    _viewModel.MaximumIndices[type] = _viewModel.MaximumIndices[type].AddOne(); 
                    // ReSharper disable once PossibleInvalidOperationException
                    AddComponent(_viewModel.PolicyProfiles, _viewModel.MaximumIndices[type].Value, names[type]);
                }

                var policyProfile = _viewModel.PolicyProfiles
                    .Where(policyView => policyView.IsVisible && policyView.UmbrellaType == null)
                    .OrderByDescending(x => x.DisplayOrder).FirstOrDefault();
                if (policyProfile != null)
                {
                    policyProfile.Sublines.Add(subline);
                    policyProfile.Sublines.Sort(x => x.ShortNameWithLob);
                }
            }

            if (subline.HasStateProfile && !ContainsSubline(_viewModel.StateProfiles, subline))
            {
                type = ComponentViewType.State;
                if (!_viewModel.StateProfiles.Any(x => x.IsVisible))
                {
                    _viewModel.MaximumIndices[type] = _viewModel.MaximumIndices[type].AddOne();
                    // ReSharper disable once PossibleInvalidOperationException
                    AddComponent(_viewModel.StateProfiles, _viewModel.MaximumIndices[type].Value, names[type]);
                }

                var state = _viewModel.StateProfiles.Where(x => x.IsVisible).OrderByDescending(x => x.DisplayOrder).First();
                state.Sublines.Add(subline);
                state.Sublines.Sort(x => x.ShortNameWithLob);
            }

            if (!ContainsSubline(_viewModel.ExposureSets, subline))
            {
                type = ComponentViewType.ExposureSet; 
                if (!_viewModel.ExposureSets.Any(x => x.IsVisible))
                {
                    _viewModel.MaximumIndices[type] = _viewModel.MaximumIndices[type].AddOne();
                    // ReSharper disable once PossibleInvalidOperationException
                    AddComponent(_viewModel.ExposureSets, _viewModel.MaximumIndices[type].Value, names[type]);
                }

                var expo = _viewModel.ExposureSets.Where(x => x.IsVisible).OrderByDescending(x => x.DisplayOrder).First();
                expo.Sublines.Add(subline);
                expo.Sublines.Sort(x => x.ShortNameWithLob);
            }

            if (!ContainsSubline(_viewModel.AggregateLossSets, subline))
            {
                if (!_viewModel.AggregateLossSets.Any(x => x.IsVisible))
                {
                    type = ComponentViewType.AggregateLossSet;
                    _viewModel.MaximumIndices[type] = _viewModel.MaximumIndices[type].AddOne();
                    // ReSharper disable once PossibleInvalidOperationException
                    AddComponent(_viewModel.AggregateLossSets, _viewModel.MaximumIndices[type].Value, names[type]);
                }

                var agg = _viewModel.AggregateLossSets.Where(x => x.IsVisible).OrderByDescending(x => x.DisplayOrder).First();
                agg.Sublines.Add(subline);
                agg.Sublines.Sort(x => x.ShortNameWithLob);
            }

            if (!ContainsSubline(_viewModel.IndividualLossSets, subline))
            {
                if (!_viewModel.IndividualLossSets.Any(x => x.IsVisible))
                {
                    type = ComponentViewType.IndividualLossSet;
                    _viewModel.MaximumIndices[type] = _viewModel.MaximumIndices[type].AddOne();
                    // ReSharper disable once PossibleInvalidOperationException
                    AddComponent(_viewModel.IndividualLossSets, _viewModel.MaximumIndices[type].Value, names[type]);
                }

                var individualLoss = _viewModel.IndividualLossSets.Where(x => x.IsVisible).OrderByDescending(x => x.DisplayOrder).First();
                individualLoss.Sublines.Add(subline);
                individualLoss.Sublines.Sort(x => x.ShortNameWithLob);
            }

            if (!ContainsSubline(_viewModel.RateChangeSets, subline))
            {
                if (!_viewModel.RateChangeSets.Any(x => x.IsVisible))
                {
                    type = ComponentViewType.RateChangeSet;
                    _viewModel.MaximumIndices[type] = _viewModel.MaximumIndices[type].AddOne();
                    // ReSharper disable once PossibleInvalidOperationException
                    AddComponent(_viewModel.RateChangeSets, _viewModel.MaximumIndices[type].Value, names[type]);
                }

                var rateChangeSets = _viewModel.RateChangeSets.Where(x => x.IsVisible).OrderByDescending(x => x.DisplayOrder).First();
                rateChangeSets.Sublines.Add(subline);
                rateChangeSets.Sublines.Sort(x => x.ShortNameWithLob);
            }

            _viewModel.OkButtonEnabled = GetOkButtonEnabledStatus();
        }

        private void Available_OnPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Available_OnMouseDoubleClick(sender, e);
        }

        private void UserPreferenceSublineFlexibilityCheckBox_OnClick(object sender, RoutedEventArgs e)
        {
            var up = UserPreferences.ReadFromFile();
            up.AllowSublineFlexibility = _viewModel.IsUserPreferenceAllowSublinesCustomization;
            up.WriteToFile();

            RedrawForm();
        }

        private string MapToLineOfBusinessName(LineOfBusinessType lineOfBusinessType)
        {
            switch (lineOfBusinessType)
            {
                case LineOfBusinessType.WorkersCompensation:
                    return _workersCompensationName;
                case LineOfBusinessType.Property:
                    return _propertyName;
                case LineOfBusinessType.PhysicalDamage:
                    return _autoPhysicalDamageName;
                case LineOfBusinessType.Liability:
                    return string.Empty;
                default:
                    return string.Empty;
            }
        }

        private static bool IsUmbrellaOkay(LineOfBusinessType lineOfBusinessType)
        {
            switch (lineOfBusinessType)
            {
                case LineOfBusinessType.WorkersCompensation:
                    return false;
                case LineOfBusinessType.Property:
                    return false;
                case LineOfBusinessType.PhysicalDamage:
                    return false;
                case LineOfBusinessType.Liability:
                    return true;
                default:
                    return false;
            }
        }

        
    }

    internal static class NullableIntegerExtensions
    {
        internal static int? AddOne(this int? number)
        {
            int? newNumber = number + 1 ?? 0;
            return newNumber;
        }
    }

    internal static class SublinesExtensions
    {
        internal static bool ContainsProperty(this IEnumerable<ISubline> sublines)
        {
            return sublines?.Any(sl => sl.LineOfBusinessType == LineOfBusinessType.Property) == true;
        }

        internal static bool ContainsCommercial(this IEnumerable<ISubline> sublines)
        {
            return sublines?.Any(sl => !sl.IsPersonal) == true;
        }

        internal static bool ContainsPolicy(this IEnumerable<ISubline> sublines)
        {
            return sublines?.Any(sl => sl.HasPolicyProfile) == true;
        }

        internal static bool ContainsState(this IEnumerable<ISubline> sublines)
        {
            return sublines?.Any(sl => sl.HasStateProfile) == true;
        }

        internal static bool ContainsSubline(this IEnumerable<ISubline> sublines, ISubline subline)
        {
            //the ordinary "contains" works in most cases
            //however, during the rebuild there are issues that I couldn't figure out
            //so just compare code to code
            return sublines != null && sublines.Contains(subline, new SublineComparer());
        }

        internal static void RemoveSubline(this IList<ISubline> sublines, ISubline subline)
        {
            if (!sublines.ContainsSubline(subline)) return;

            //the ordinary "remove" works in most cases
            //however, during the rebuild there are issues that I couldn't figure out
            //so just compare code to code
            var index = 0;
            foreach (var sublineItem in sublines)
            {
                if (sublineItem.Code == subline.Code) break;
                index++;
            }
            
            sublines.RemoveAt(index);
        }

        internal static bool IsProperty(this ISubline subline)
        {
            return subline.LineOfBusinessType == LineOfBusinessType.Property;
        }
    }
}