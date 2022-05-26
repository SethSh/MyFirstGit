using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class SublineWizardManager
    {
        public void ModifySublinesWithWizard(IWorkbookLogger logger)
        {
            try
            {
                ModifySublinesWithWizard();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Modify {BexConstants.SublineName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static SublineSelectorWizardResponse LaunchSublineSelectorWizard(ISegment segment)
        {
            var title = $"{BexConstants.SegmentName} {BexConstants.SublineName} Wizard";

            var viewModel = new SublineWizardViewModel(segment);
            var sublineWizard = new SublineWizard(viewModel);
            var form = new SublineWizardForm(sublineWizard)
            {
                Text = title,
                Height = (int) FormSizeHeight.ExtraLarge,
                Width = (int) FormSizeWidth.ExtraLarge,
                StartPosition = FormStartPosition.CenterScreen
            };
            form.ShowDialog();

            return new SublineSelectorWizardResponse {FormResponse = sublineWizard.Response, ViewModel = viewModel};
        }

        public static void MapSublineWizardViewModel(ISegment segment, ISublineSelectorWizardViewModel viewModel)
        {
            segment.UpdateSublines(viewModel.SegmentSublines);

            UpdatePolicyProfiles(segment, viewModel.PolicyProfiles.Where(viewModelProfile => viewModelProfile.IsVisible).ToList());
            UpdateStateProfiles(segment, viewModel.StateProfiles.Where(viewModelProfile => viewModelProfile.IsVisible).ToList());

            RetroFitPropertyProfiles(segment);
            
            RetroFitWorkersCompProfiles(segment);
            UpdateWorkersCompMinnesotaRetention(segment, viewModel.SegmentSublines);
            UpdateWorkersCompStateHazardProfile(segment, viewModel.SegmentSublines);
            UpdateWorkersCompClassCodeProfile(segment, viewModel.SegmentSublines);
            UpdateWorkersCompStateAttachmentProfile(segment, viewModel.SegmentSublines);

            UpdateExposureSets(segment, viewModel.ExposureSets.Where(viewModelProfile => viewModelProfile.IsVisible).ToList());
            UpdateAggregateLossSets(segment, viewModel.AggregateLossSets.Where(viewModelProfile => viewModelProfile.IsVisible).ToList());
            UpdateIndividualLossSets(segment, viewModel.IndividualLossSets.Where(viewModelProfile => viewModelProfile.IsVisible).ToList());
            UpdateRateChangeSets(segment, viewModel.RateChangeSets.Where(viewModelProfile => viewModelProfile.IsVisible).ToList());

            ReflectChangedSublines(segment, viewModel);
        }

        private static void ModifySublinesWithWizard()
        {
            var rangeValidator = new SegmentWorksheetValidator();
            if (!rangeValidator.Validate()) return;

            var segment = rangeValidator.Segment;

            if (!segment.IsStructureModifiable)
            {
                var message = segment.GetBlockModificationsMessage(BexConstants.SublineName);
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            if (!segment.IsSelected)
            {
                var message = $"Modifying {BexConstants.SublineName.ToLower()}s requires first selecting a " +
                              $"{BexConstants.SegmentName.ToLower()} node in the {BexConstants.InventoryTreeName.ToLower()}.";
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            var sublineSelectorResponse = LaunchSublineSelectorWizard(segment);
            if (sublineSelectorResponse.FormResponse == FormResponse.Cancel) return;

            MapSublineWizardViewModel(segment, sublineSelectorResponse.ViewModel);
            if (segment.IsUmbrella && segment.ContainsAnyCommercialSublines && !segment.UmbrellaExcelMatrix.HeaderRangeName.ExistsInWorkbook())
            {
                var umbrellaSelector = new UmbrellaTypeAllocatorManager();
                umbrellaSelector.SelectUmbrellaTypes(new StackTraceLogger());
            }

            segment.WorksheetManager.ModifyWorksheetWrapper();

            if (segment.IsUmbrella)
            {
                ExcelSheetActivateEventManager.RefreshUmbrellaWizardButton(segment);
            }

            ExcelSheetActivateEventManager.RefreshStateSortOptions(segment);
            ExcelSheetActivateEventManager.RefreshLineOfBusinessMenus(segment);
        }

        
        
        private static void RetroFitPropertyProfiles(ISegment segment)
        {
            if (!segment.IsProperty) return;

            foreach (var propertySubline in segment.Where(subline => subline.IsProperty()))
            {
                if (!segment.ConstructionTypeProfiles.Any()) segment.AddConstructionTypeProfile(propertySubline);
                if (!segment.ProtectionClassProfiles.Any()) segment.AddProtectionClassProfile(propertySubline);
                if (!segment.TotalInsuredValueProfiles.Any()) segment.AddTotalInsuredValueProfile(propertySubline);
                if (!segment.OccupancyTypeProfiles.Any()) segment.AddOccupancyTypeProfile(propertySubline);
            }
        }

        private static void RetroFitWorkersCompProfiles(ISegment segment)
        {
            if (!segment.IsWorkersComp) return;
            
            if (segment.WorkersCompClassCodeProfile == null) segment.AddWorkersCompClassCodeProfile();
            if (segment.WorkersCompStateAttachmentProfile == null) segment.AddWorkersCompStateAttachmentProfile();
            if (segment.WorkersCompStateHazardGroupProfile == null) segment.AddWorkersCompStateHazardGroupProfile();
            if (segment.WorkersCompMinnesotaRetention == null) segment.AddWorkersCompMinnesotaRetention();
        }



        private static void ReflectChangedSublines(ISegment segment, ISublineSelectorWizardViewModel viewModel)
        {
            foreach (var item in segment.PolicyProfiles)
            {
                var sublines = viewModel.PolicyProfiles.Single(x => x.Guid == item.Guid).Sublines;
                var policyExcelMatrix = (PolicyExcelMatrix)item.CommonExcelMatrix;
                if (!policyExcelMatrix.IsNotEqualTo(sublines)) continue;

                policyExcelMatrix.UpdateSublines(sublines);
                item.IsDirty = true;
            }

            foreach (var item in segment.StateProfiles)
            {
                var changeSublines = viewModel.StateProfiles.Single(x => x.Guid == item.Guid).Sublines;
                if (!item.ExcelMatrix.IsNotEqualTo(changeSublines)) continue;

                item.ExcelMatrix.UpdateSublines(changeSublines);
                item.IsDirty = true;
            }
            
            
            
            foreach (var item in segment.ExposureSets)
            {
                var changeSublines = viewModel.ExposureSets.Single(x => x.Guid == item.Guid).Sublines;
                if (!item.ExcelMatrix.IsNotEqualTo(changeSublines)) continue;

                item.ExcelMatrix.UpdateSublines(changeSublines);
                item.IsDirty = true;
            }

            foreach (var item in segment.AggregateLossSets)
            {
                var changeSublines = viewModel.AggregateLossSets.Single(x => x.Guid == item.Guid).Sublines;
                if (!item.ExcelMatrix.IsNotEqualTo(changeSublines)) continue;

                item.ExcelMatrix.UpdateSublines(changeSublines);
                item.IsDirty = true;
            }

            foreach (var item in segment.IndividualLossSets)
            {
                var changeSublines = viewModel.IndividualLossSets.Single(x => x.Guid == item.Guid).Sublines;
                if (!item.ExcelMatrix.IsNotEqualTo(changeSublines)) continue;

                item.ExcelMatrix.UpdateSublines(changeSublines);
                item.IsDirty = true;
            }

            foreach (var item in segment.RateChangeSets)
            {
                var changeSublines = viewModel.RateChangeSets.Single(x => x.Guid == item.Guid).Sublines;
                if (!item.ExcelMatrix.IsNotEqualTo(changeSublines)) continue;

                item.ExcelMatrix.UpdateSublines(changeSublines);
                item.IsDirty = true;
            }
        }

        private static void UpdatePolicyProfiles(ISegment segment, IList<ComponentView> viewModelProfiles)
        {
            var isInitialMap = !segment.PolicyProfiles.Any();
            foreach (var viewModelProfile in viewModelProfiles)
            {
                var guid = viewModelProfile.Guid;
                var id = viewModelProfile.Id;
                var displayOrder = viewModelProfile.DisplayOrder;
                
                if (isInitialMap || segment.PolicyProfiles.SingleOrDefault(x => x.Guid == guid) == null)
                {
                    var profile = new PolicyProfile(segment.Id, id) {Guid = guid};
                    profile.ExcelMatrix.IntraDisplayOrder = viewModelProfile.DisplayOrder;
                    segment.ExcelComponents.Add(profile);
                }
                else
                {
                    var existingProfile = segment.PolicyProfiles.SingleOrDefault(y => y.Guid == guid);
                    if (existingProfile == null) continue;
                    
                    var policyExcelMatrix = existingProfile.ExcelMatrix;
                    if (policyExcelMatrix.IntraDisplayOrder != displayOrder)
                    {
                        policyExcelMatrix.IntraDisplayOrder = displayOrder;
                    }
                }
            }

            var viewModelProfileGuids = viewModelProfiles.Select(viewModel => viewModel.Guid).ToList();
            segment.ExcelComponents.RemoveAll(ec => ec is PolicyProfile && !viewModelProfileGuids.Contains(ec.Guid));
        }

        private static void UpdateStateProfiles(ISegment segment, IList<ComponentView> viewModelProfiles)
        {
            var isInitialMap = !segment.StateProfiles.Any();
            viewModelProfiles.ForEach(viewModelProfile =>
            {
                var guid = viewModelProfile.Guid;
                var id = viewModelProfile.Id;
                var displayOrder = viewModelProfile.DisplayOrder;
                
                if (isInitialMap || segment.StateProfiles.SingleOrDefault(y => y.Guid == guid) == null)
                {
                    var profile = new StateProfile(segment.Id, id) {Guid = guid};
                    profile.ExcelMatrix.IntraDisplayOrder = viewModelProfile.DisplayOrder;
                    segment.ExcelComponents.Add(profile);
                }
                else
                {
                    var existingProfile = segment.StateProfiles.SingleOrDefault(y => y.Guid == guid);
                    if (existingProfile == null) return;
                    
                    if (existingProfile.IntraDisplayOrder != displayOrder)
                    {
                        existingProfile.ExcelMatrix.IntraDisplayOrder = displayOrder;
                    }
                }
            });

            var viewModelProfileGuids = viewModelProfiles.Select(viewModel => viewModel.Guid).ToList();
            segment.ExcelComponents.RemoveAll(ec => ec is StateProfile && !viewModelProfileGuids.Contains(ec.Guid));
        }

        
        private static void UpdateWorkersCompMinnesotaRetention(ISegment segment, IList<ISubline> changedSublines)
        {
            if (changedSublines.All(subline => subline.IsWorkersComp))
            {
                if (segment.WorkersCompMinnesotaRetention == null)
                {
                    segment.ExcelComponents.Add(new MinnesotaRetention(segment.Id));
                }
                return;
            }

            segment.ExcelComponents.RemoveAll(ec => ec is MinnesotaRetention);

        }
        
        private static void UpdateWorkersCompStateHazardProfile(ISegment segment, IList<ISubline> changedSublines)
        {
            if (changedSublines.All(subline => subline.IsWorkersComp))
            {
                if (segment.WorkersCompStateHazardGroupProfile == null)
                {
                    var profile = new WorkersCompStateHazardGroupProfile(segment.Id);
                    profile.ExcelMatrix.IsIndependent = !segment.IsWorkerCompClassCodeActive &&
                                                UserPreferences.ReadFromFile().WorkersCompStateByHazardGroupIsIndependent;
                    segment.ExcelComponents.Add(profile);
                }
                return;
            }

            segment.ExcelComponents.RemoveAll(ec => ec is WorkersCompStateHazardGroupProfile);
        }

        private static void UpdateWorkersCompClassCodeProfile(ISegment segment, IList<ISubline> changedSublines)
        {
            if (changedSublines.All(subline => subline.IsWorkersComp))
            {
                if (segment.WorkersCompClassCodeProfile == null && segment.IsWorkerCompClassCodeActive)
                {
                    segment.ExcelComponents.Add(new WorkersCompClassCodeProfile(segment.Id));
                }
                return;
            }

            segment.ExcelComponents.RemoveAll(ec => ec is WorkersCompClassCodeProfile);
        }

        private static void UpdateWorkersCompStateAttachmentProfile(ISegment segment, IList<ISubline> changedSublines)
        {
            if (changedSublines.All(subline => subline.IsWorkersComp))
            {
                if (segment.WorkersCompStateAttachmentProfile == null)
                {
                    segment.ExcelComponents.Add(new WorkersCompStateAttachmentProfile(segment.Id));
                }
                return;
            }

            segment.ExcelComponents.RemoveAll(ec => ec is WorkersCompStateAttachmentProfile);
        }
        
        
        
        private static void UpdateExposureSets(ISegment segment, IList<ComponentView> viewModelHistoricals)
        {
            var isInitialMap = !segment.ExposureSets.Any();
            foreach (var viewModelHistorical in viewModelHistoricals)
            {
                var guid = viewModelHistorical.Guid;
                var id = viewModelHistorical.Id;
                var displayOrder = viewModelHistorical.DisplayOrder;
                
                if (isInitialMap || !segment.ExposureSets.ContainsGuid(guid))
                {
                    var exposureSet = new ExposureSet(segment.Id, id) {Guid = guid};
                    exposureSet.ExcelMatrix.IntraDisplayOrder = viewModelHistorical.DisplayOrder;
                    segment.ExcelComponents.Add(exposureSet);
                }
                else
                {
                    var existingExposureSet = segment.ExposureSets.Get(guid);
                    if (existingExposureSet == null) continue;
                    
                    if (existingExposureSet.IntraDisplayOrder != displayOrder)
                    {
                        existingExposureSet.CommonExcelMatrix.IntraDisplayOrder = displayOrder;
                    }
                }
            }

            var viewModelHistoricalGuids = viewModelHistoricals.Select(viewModel => viewModel.Guid).ToList();
            segment.ExcelComponents.RemoveAll(ec => ec is ExposureSet && !viewModelHistoricalGuids.Contains(ec.Guid));
        }

        private static void UpdateAggregateLossSets(ISegment segment, IList<ComponentView> viewModelHistoricals)
        {
            var isInitialMap = !segment.AggregateLossSets.Any();
            viewModelHistoricals.ForEach(viewModelHistorical =>
            {
                var guid = viewModelHistorical.Guid;
                var id = viewModelHistorical.Id;
                var displayOrder = viewModelHistorical.DisplayOrder;
                
                if (isInitialMap || segment.AggregateLossSets.SingleOrDefault(y => y.Guid == guid) == null)
                {
                    var lossSet = new AggregateLossSet(segment.Id, id) {Guid = guid};
                    lossSet.ExcelMatrix.IntraDisplayOrder = viewModelHistorical.DisplayOrder;
                    segment.ExcelComponents.Add(lossSet);
                }
                else
                {
                    var existingLossSet = segment.AggregateLossSets.SingleOrDefault(y => y.Guid == guid);
                    if (existingLossSet == null) return;
                    
                    if (existingLossSet.IntraDisplayOrder != displayOrder)
                    {
                        existingLossSet.ExcelMatrix.IntraDisplayOrder = displayOrder;
                    }
                }
            });

            var viewModelAggGuids = viewModelHistoricals.Select(viewModel => viewModel.Guid).ToList();
            segment.ExcelComponents.RemoveAll(ec => ec is AggregateLossSet && !viewModelAggGuids.Contains(ec.Guid));
        }

        private static void UpdateIndividualLossSets(ISegment segment, IList<ComponentView> viewModelHistoricals)
        {
            var isInitialMap = !segment.IndividualLossSets.Any();
            viewModelHistoricals.ForEach(viewModelHistorical =>
            {
                var guid = viewModelHistorical.Guid;
                var id = viewModelHistorical.Id;
                var displayOrder = viewModelHistorical.DisplayOrder;
                
                if (isInitialMap || segment.IndividualLossSets.SingleOrDefault(y => y.Guid == guid) == null)
                {
                    var lossSet = new IndividualLossSet(segment.Id, id) {Guid = guid};
                    lossSet.ExcelMatrix.IntraDisplayOrder = viewModelHistorical.DisplayOrder;
                    segment.ExcelComponents.Add(lossSet);
                }
                else
                {
                    var existingLossSet = segment.IndividualLossSets.SingleOrDefault(y => y.Guid == guid);
                    if (existingLossSet == null) return;

                    if (existingLossSet.IntraDisplayOrder != displayOrder)
                    {
                        existingLossSet.ExcelMatrix.IntraDisplayOrder = displayOrder;
                    }
                }
            });

            var viewModelHistoricalGuids = viewModelHistoricals.Select(viewModel => viewModel.Guid).ToList();
            segment.ExcelComponents.RemoveAll(ec => ec is IndividualLossSet && !viewModelHistoricalGuids.Contains(ec.Guid));
        }


        private static void UpdateRateChangeSets(ISegment segment, IList<ComponentView> viewModels)
        {
            var isInitialMap = !segment.RateChangeSets.Any();
            viewModels.ForEach(viewModelHistorical =>
            {
                var guid = viewModelHistorical.Guid;
                var id = viewModelHistorical.Id;
                var displayOrder = viewModelHistorical.DisplayOrder;
                
                if (isInitialMap || segment.RateChangeSets.SingleOrDefault(y => y.Guid == guid) == null)
                {
                    var rateChangeSet = new RateChangeSet(segment.Id, id) {Guid = guid};
                    rateChangeSet.ExcelMatrix.IntraDisplayOrder = viewModelHistorical.DisplayOrder;
                    segment.ExcelComponents.Add(rateChangeSet);
                }
                else
                {
                    var rateChangeSet = segment.RateChangeSets.SingleOrDefault(y => y.Guid == guid);
                    if (rateChangeSet == null) return;

                    if (rateChangeSet.IntraDisplayOrder != displayOrder)
                    {
                        rateChangeSet.ExcelMatrix.IntraDisplayOrder = displayOrder;
                    }
                }
            });

            var viewModelHistoricalGuids = viewModels.Select(viewModel => viewModel.Guid).ToList();
            segment.ExcelComponents.RemoveAll(ec => ec is RateChangeSet && !viewModelHistoricalGuids.Contains(ec.Guid));
        }

        public class SublineSelectorWizardResponse
        {
            public FormResponse FormResponse { get; set; }
            public SublineWizardViewModel ViewModel { get; set; }
        }
    }
}
