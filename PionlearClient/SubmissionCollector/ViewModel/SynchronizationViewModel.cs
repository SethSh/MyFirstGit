using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using PionlearClient;
using PionlearClient.CollectorClientPlus;
using PionlearClient.CollectorClientPlus.Extensions;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.Extensions;
using PionlearClient.Model;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ViewModel
{
    public class SynchronizationViewModel : BaseSynchronizationViewModel
    {
        private readonly Package _package;
        private readonly BexCommunicationManager _bexCommunicationManager;
        private SynchronizationView _packageView;
        private PackageModel _packageModel;

        public SynchronizationViewModel(Package package, BexCommunicationManager bexCommunicationManager)
        {
            _package = package;
            _bexCommunicationManager = bexCommunicationManager;

            LoadingRowPixels = LoadingRowHeight;
            StatusMessage = "Loading ...";
            ValidationRowPixels = ErrorRowPixels = LegendRowPixels = BottomRowPixels = new GridLength(0);
            HeadingRowPixels = HeadingRowHeight;
            BodyRowPixels = BodyRowHeight;
            ValidationMessage = ErrorMessage = string.Empty;
        }

        public void CreateViews()
        {
            try
            {
                StatusMessage = $"Loading {BexConstants.UploadPreviewName.ToLower()} ...";

                if (!Validate())
                {
                    ValidationMessage = "Unable to preview: validation failed";
                    ValidationRowPixels = ValidationRowHeight;
                    SetCommonFailures();
                    return;
                }

                CreatePackageView();

                HeadingRowPixels = HeadingRowHeight;
                LoadingRowPixels = ValidationRowPixels = ErrorRowPixels = new GridLength(0);
                BodyRowPixels = BodyRowHeight;
                LegendRowPixels = LegendRowHeight;
                BottomRowPixels = BottomRowHeight;
                ShowZoom = true;
            }
            catch (Exception e)
            {
                ErrorMessage = $"Unable to preview: {e.Message}";
                ErrorRowPixels = ErrorRowHeight;
                SetCommonFailures();
            }
        }

        public void AddPackage()
        {
            PackageSynchronizationViews.Add(_packageView);
        }

        public override bool IsEntirePackageSynchronized
        {
            get
            {
                if (_packageView?.SynchronizationCode != SynchronizationCode.InSynchronization) return false;

                var segmentViews = _packageView.ChildViews;
                if (segmentViews.Any(sv => sv.SynchronizationCode != SynchronizationCode.InSynchronization)) return false;

                foreach (var segmentView in segmentViews)
                {
                    if (segmentView.ChildViews.Any(cv => cv.SynchronizationCode != SynchronizationCode.InSynchronization)) return false;
                }

                return true;
            }
        }

        private void SetCommonFailures()
        {
            HeadingRowPixels = LoadingRowPixels = BodyRowPixels = LegendRowPixels = new GridLength(0);
            BodyRowPixels = BodyRowHeight;
            BottomRowPixels = BottomRowHeight;
            ShowZoom = false;
        }

        public bool Validate()
        {
            _packageModel = _package.CreatePackageModel();
            if (_package.ExcelValidation.Length > 0) return false;

            var validation = _packageModel.Validate();
            return validation.Length <= 0;
        }

        public void CreatePackageView()
        {
            var synchronizationCode = GetPackageSynchronizationCode();

            _packageView = new SynchronizationView {Name = _package.Name, SynchronizationCode = synchronizationCode};

            //segments the workbook knows about
            var serverSegments = _bexCommunicationManager.GetServerSegments(_package.SourceId);
            foreach (var segment in _package.Segments)
            {
                _packageView.ChildViews.Add(CreateSegmentView(segment, serverSegments));
            }

            //segments the workbook doesn't know about
            var localSegmentIds = _package.Segments.Where(seg => seg.SourceId.HasValue).Select(seg => seg.SourceId.Value);
            var serverOrphanSegments = serverSegments.Where(serverSegment => !localSegmentIds.Contains(serverSegment.Key));
            foreach (var serverOrphanSegment in serverOrphanSegments)
            {
                _packageView.ChildViews.Add(CreateServerOnlyView(serverOrphanSegment.Value.Name));
            }
        }

        private SynchronizationCode GetPackageSynchronizationCode()
        {
            SynchronizationCode synchronizationCode;
            if (!_packageModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(_packageModel.SourceId.HasValue, "Package Id is null");
                var packageId = _packageModel.SourceId.Value;
                var serverPackageModel = _bexCommunicationManager.TryToGetSubmissionPackageModel(packageId);

                if (serverPackageModel == null)
                {
                    synchronizationCode = SynchronizationCode.New;
                }
                else
                {
                    var collectorPackageModel = _packageModel.Map();
                    synchronizationCode = collectorPackageModel.IsEqualTo(serverPackageModel)
                        ? SynchronizationCode.InSynchronization
                        : SynchronizationCode.NotInSynchronization;
                }
            }

            return synchronizationCode;
        }

        private SynchronizationView CreateSegmentView(ISegment segment,
            Dictionary<long, CollectorApi.SubmissionSegmentModel> serverSegments)
        {
            var segmentModel = _packageModel.SegmentModels.Single(seg => seg.Guid == segment.Guid);
            var segmentSynchronizationCode = GetSegmentSynchronizationCode(segmentModel, serverSegments);

            var segmentView = new SynchronizationView {Name = segment.Name, SynchronizationCode = segmentSynchronizationCode};

            CreateHazard(segment, segmentModel, segmentView);
            CreatePolicy(segment, segmentModel, segmentView);
            CreateState(segment, segmentModel, segmentView);

            CreateProtectionClass(segment, segmentModel, segmentView);
            CreateConstructionType(segment, segmentModel, segmentView);
            CreateOccupancyType(segment, segmentModel, segmentView);
            CreateTotalInsuredValue(segment, segmentModel, segmentView);

            CreateWorkersCompStateAttachment(segment, segmentModel, segmentView);
            CreateWorkersCompStateHazard(segment, segmentModel, segmentView);
            CreateWorkersCompClassCode(segment, segmentModel, segmentView);
            
            CreateExposureSet(segment, segmentModel, segmentView);
            CreateAggregateLossSet(segment, segmentModel, segmentView);
            CreateIndividualLossSet(segment, segmentModel, segmentView);
            CreateRateChangeSet(segment, segmentModel, segmentView);

            return segmentView;
        }

        private static SynchronizationCode GetSegmentSynchronizationCode(SegmentModel segmentModel,
            Dictionary<long, CollectorApi.SubmissionSegmentModel> serverSegments)
        {
            SynchronizationCode synchronizationCode;
            if (!segmentModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(segmentModel.SourceId.HasValue, "Segment Id is null");
                var segmentId = segmentModel.SourceId.Value;
                if (serverSegments.ContainsKey(segmentId))
                {
                    var collectorSegmentModel = segmentModel.Map();
                    synchronizationCode = collectorSegmentModel.IsEqualTo(serverSegments[segmentId])
                        ? SynchronizationCode.InSynchronization
                        : SynchronizationCode.NotInSynchronization;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private void CreateHazard(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //hazard the workbook knows about
            var serverHazardProfiles = _bexCommunicationManager.GetServerHazardProfiles(_package.SourceId, segment.SourceId);
            var profilesToHide = segment.HazardProfiles.Where(hp => !hp.SourceId.HasValue && !hp.ExcelMatrix.Items.Any());
            segment.HazardProfiles.OrderBy(hp => hp.IntraDisplayOrder)
                .Except(profilesToHide)
                .ForEach(hp => segmentView.ChildViews.Add(CreateHazardProfileView(segmentModel, hp, serverHazardProfiles)));

            //hazard the workbook doesn't know about
            var localIds = segment.HazardProfiles.Where(hp => hp.SourceId.HasValue).Select(hp => hp.SourceId.Value);
            var serverOrphans = serverHazardProfiles.Where(serv => !localIds.Contains(serv.Key));
            serverOrphans.ForEach(serverOrphan => segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.DistributionName)));
        }

        private static SynchronizationView CreateHazardProfileView(SegmentModel segmentModel, HazardProfile hazardProfile,
            Dictionary<long, CollectorApi.HazardDistributionModel> serverProfiles)
        {
            var hazardModel = segmentModel.HazardModels.Single(hm => hm.Guid == hazardProfile.Guid);
            var hazardSynchronizationCode = GetHazardSynchronizationCode(hazardModel, serverProfiles);

            var hazardProfileView = new SynchronizationView
            {
                Name = hazardProfile.CommonExcelMatrix.FullName, SynchronizationCode = hazardSynchronizationCode
            };

            return hazardProfileView;
        }

        private static SynchronizationCode GetHazardSynchronizationCode(HazardModel hazardModel,
            Dictionary<long, CollectorApi.HazardDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!hazardModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(hazardModel.SourceId.HasValue, "Hazard Id is null");
                var hazardId = hazardModel.SourceId.Value;
                if (serverProfiles.ContainsKey(hazardId))
                {
                    var collectorHazardModel = hazardModel.Map();
                    synchronizationCode = collectorHazardModel.IsEqualTo(serverProfiles[hazardId]) ? SynchronizationCode.InSynchronization :
                        hazardModel.Items.Any() ? SynchronizationCode.NotInSynchronization : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private void CreatePolicy(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //policy the workbook knows about
            var serverPolicyProfiles = _bexCommunicationManager.GetServerPolicyProfiles(_package.SourceId, segment.SourceId);
            var profilesToHide = segment.PolicyProfiles.Where(hp => !hp.SourceId.HasValue && !hp.ExcelMatrix.Items.Any());
            segment.PolicyProfiles.OrderBy(pp => pp.IntraDisplayOrder)
                .Except(profilesToHide)
                .ForEach(pp => segmentView.ChildViews.Add(CreatePolicyProfileView(segmentModel, pp, serverPolicyProfiles)));

            //policy the workbook doesn't know about
            var localIds = segment.PolicyProfiles.Where(hp => hp.SourceId.HasValue).Select(hp => hp.SourceId.Value);
            var serverOrphans = serverPolicyProfiles.Where(serverProfile => !localIds.Contains(serverProfile.Key));
            serverOrphans.ForEach(serverOrphan => segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.DistributionName)));
        }

        private static SynchronizationView CreatePolicyProfileView(SegmentModel segmentModel, PolicyProfile policyProfile,
            Dictionary<long, CollectorApi.PolicyDistributionModel> serverProfiles)
        {
            var policyModel = segmentModel.PolicyModels.Single(hm => hm.Guid == policyProfile.Guid);
            var policySynchronizationCode = GetPolicySynchronizationCode(policyModel, serverProfiles);

            var policyProfileView = new SynchronizationView
            {
                Name = policyProfile.CommonExcelMatrix.FullName, SynchronizationCode = policySynchronizationCode
            };

            return policyProfileView;
        }

        private static SynchronizationCode GetPolicySynchronizationCode(PolicyModel policyModel,
            Dictionary<long, CollectorApi.PolicyDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!policyModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(policyModel.SourceId.HasValue, "Policy Id is null");
                var policyId = policyModel.SourceId.Value;
                if (serverProfiles.ContainsKey(policyId))
                {
                    var collectorPolicyModel = policyModel.Map();
                    synchronizationCode = collectorPolicyModel.IsEqualTo(serverProfiles[policyId]) ? SynchronizationCode.InSynchronization :
                        policyModel.Items.Any() ? SynchronizationCode.NotInSynchronization : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private void CreateState(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //state the workbook knows about
            var serverStateProfiles = _bexCommunicationManager.GetServerStateProfiles(_package.SourceId, segment.SourceId);
            var profilesToHide = segment.StateProfiles.Where(hp => !hp.SourceId.HasValue && !hp.ExcelMatrix.Items.Any());
            segment.StateProfiles.OrderBy(sp => sp.IntraDisplayOrder)
                .Except(profilesToHide)
                .ForEach(sp => segmentView.ChildViews.Add(CreateStateProfileView(segmentModel, sp, serverStateProfiles)));

            //state the workbook doesn't know about
            var localIds = segment.StateProfiles.Where(hp => hp.SourceId.HasValue).Select(hp => hp.SourceId.Value);
            var serverOrphans = serverStateProfiles.Where(serverProfile => !localIds.Contains(serverProfile.Key));
            serverOrphans.ForEach(serverOrphanState =>
                segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphanState.Value.DistributionName)));
        }

        private static SynchronizationView CreateStateProfileView(SegmentModel segmentModel, StateProfile stateProfile,
            Dictionary<long, CollectorApi.StateDistributionModel> serverProfiles)
        {
            var stateModel = segmentModel.StateModels.Single(hm => hm.Guid == stateProfile.Guid);
            var stateSynchronizationCode = GetStateSynchronizationCode(stateModel, serverProfiles);

            var stateProfileView = new SynchronizationView
            {
                Name = stateProfile.CommonExcelMatrix.FullName, SynchronizationCode = stateSynchronizationCode
            };

            return stateProfileView;
        }

        private void CreateConstructionType(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //profiles the workbook knows about
            var serverProfiles = _bexCommunicationManager.GetServerConstructionTypeProfiles(_package.SourceId, segment.SourceId);
            var profilesToHide =
                segment.ConstructionTypeProfiles.Where(profile => !profile.SourceId.HasValue && !profile.ExcelMatrix.Items.Any());
            foreach (var profile in segment.ConstructionTypeProfiles.OrderBy(profile => profile.IntraDisplayOrder).Except(profilesToHide))
            {
                segmentView.ChildViews.Add(CreateConstructionTypeProfileView(segmentModel, profile, serverProfiles));
            }

            //profiles the workbook doesn't know about
            var localIds = segment.ConstructionTypeProfiles.Where(profile => profile.SourceId.HasValue)
                .Select(profile => profile.SourceId.Value);
            var serverOrphans = serverProfiles.Where(serverProfile => !localIds.Contains(serverProfile.Key));
            foreach (var serverOrphan in serverOrphans)
            {
                segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.DistributionName));
            }
        }

        private static SynchronizationView CreateConstructionTypeProfileView(SegmentModel segmentModel,
            ConstructionTypeProfile constructionTypeProfile,
            Dictionary<long, CollectorApi.ConstructionTypeDistributionModel> serverProfiles)
        {
            var constructionTypeModel = segmentModel.ConstructionTypeModels.Single(ctm => ctm.Guid == constructionTypeProfile.Guid);
            var code = GetConstructionTypeSynchronizationCode(constructionTypeModel, serverProfiles);

            var view = new SynchronizationView {Name = constructionTypeProfile.CommonExcelMatrix.FullName, SynchronizationCode = code};

            return view;
        }

        private void CreateOccupancyType(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //profiles the workbook knows about
            var serverProfiles = _bexCommunicationManager.GetServerOccupancyTypeProfiles(_package.SourceId, segment.SourceId);
            var profilesToHide =
                segment.OccupancyTypeProfiles.Where(profile => !profile.SourceId.HasValue && !profile.ExcelMatrix.Items.Any());
            foreach (var profile in segment.OccupancyTypeProfiles.OrderBy(profile => profile.IntraDisplayOrder).Except(profilesToHide))
            {
                segmentView.ChildViews.Add(CreateOccupancyTypeProfileView(segmentModel, profile, serverProfiles));
            }

            //profiles the workbook doesn't know about
            var localIds = segment.OccupancyTypeProfiles.Where(profile => profile.SourceId.HasValue)
                .Select(profile => profile.SourceId.Value);
            var serverOrphans = serverProfiles.Where(serverProfile => !localIds.Contains(serverProfile.Key));
            foreach (var serverOrphan in serverOrphans)
            {
                segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.DistributionName));
            }
        }

        private static SynchronizationView CreateOccupancyTypeProfileView(SegmentModel segmentModel,
            OccupancyTypeProfile constructionTypeProfile, Dictionary<long, CollectorApi.OccupancyTypeDistributionModel> serverProfiles)
        {
            var model = segmentModel.OccupancyTypeModels.Single(ctm => ctm.Guid == constructionTypeProfile.Guid);
            var code = GetOccupancyTypeSynchronizationCode(model, serverProfiles);

            var view = new SynchronizationView {Name = constructionTypeProfile.CommonExcelMatrix.FullName, SynchronizationCode = code};

            return view;
        }

        private void CreateProtectionClass(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //profiles the workbook knows about
            var serverProfiles = _bexCommunicationManager.GetServerProtectionClassProfiles(_package.SourceId, segment.SourceId);
            var profilesToHide =
                segment.ProtectionClassProfiles.Where(profile => !profile.SourceId.HasValue && !profile.ExcelMatrix.Items.Any());
            foreach (var profile in segment.ProtectionClassProfiles.OrderBy(profile => profile.IntraDisplayOrder).Except(profilesToHide))
            {
                segmentView.ChildViews.Add(CreateProtectionClassProfileView(segmentModel, profile, serverProfiles));
            }

            //profiles the workbook doesn't know about
            var localIds = segment.ProtectionClassProfiles.Where(profile => profile.SourceId.HasValue)
                .Select(profile => profile.SourceId.Value);
            var serverOrphans = serverProfiles.Where(serverProfile => !localIds.Contains(serverProfile.Key));
            foreach (var serverOrphan in serverOrphans)
            {
                segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.DistributionName));
            }
        }

        private static SynchronizationView CreateProtectionClassProfileView(SegmentModel segmentModel,
            ProtectionClassProfile constructionTypeProfile, Dictionary<long, CollectorApi.ProtectionClassDistributionModel> serverProfiles)
        {
            var model = segmentModel.ProtectionClassModels.Single(ctm => ctm.Guid == constructionTypeProfile.Guid);
            var protectionClassSynchronizationCode = GetProtectionClassSynchronizationCode(model, serverProfiles);

            var profileView = new SynchronizationView
            {
                Name = constructionTypeProfile.CommonExcelMatrix.FullName, SynchronizationCode = protectionClassSynchronizationCode
            };

            return profileView;
        }

        private void CreateTotalInsuredValue(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //profiles the workbook knows about
            var serverProfiles = _bexCommunicationManager.GetServerTotalInsuredValueProfiles(_package.SourceId, segment.SourceId);
            var profilesToHide = segment.TotalInsuredValueProfiles.Where(profile => !profile.SourceId.HasValue && !profile.ExcelMatrix.Items.Any());
            foreach (var profile in segment.TotalInsuredValueProfiles.OrderBy(profile => profile.IntraDisplayOrder).Except(profilesToHide))
            {
                segmentView.ChildViews.Add(CreateTotalInsuredValueProfileView(segmentModel, profile, serverProfiles));
            }

            //profiles the workbook doesn't know about
            var localIds = segment.TotalInsuredValueProfiles.Where(profile => profile.SourceId.HasValue).Select(profile => profile.SourceId.Value);
            var serverOrphans = serverProfiles.Where(serverProfile => !localIds.Contains(serverProfile.Key));
            foreach (var serverOrphan in serverOrphans)
            {
                segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.DistributionName));
            }
        }

        private static SynchronizationView CreateTotalInsuredValueProfileView(SegmentModel segmentModel, TotalInsuredValueProfile totalInsuredValueProfile,
            Dictionary<long, CollectorApi.TotalInsuredValueDistributionModel> serverProfiles)
        {
            var model = segmentModel.TotalInsuredValueModels.Single(ctm => ctm.Guid == totalInsuredValueProfile.Guid);
            var code = GetTotalInsuredValueSynchronizationCode(model, serverProfiles);

            var profileView = new SynchronizationView
            {
                Name = totalInsuredValueProfile.CommonExcelMatrix.FullName,
                SynchronizationCode = code
            };

            return profileView;
        }



        private void CreateWorkersCompStateAttachment(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //profiles the workbook knows about
            var serverProfiles = _bexCommunicationManager.GetServerWorkersCompStateAttachmentProfiles(_package.SourceId, segment.SourceId);
            if (segment.IsWorkersComp)
            {
                var showProfile = segment.WorkersCompStateAttachmentProfile.SourceId.HasValue ||
                                  segment.WorkersCompStateAttachmentProfile.ExcelMatrix.Items.Any();
                if (showProfile)
                {
                    segmentView.ChildViews.Add(CreateWorkersCompStateAttachmentProfileView(segmentModel,
                        segment.WorkersCompStateAttachmentProfile, serverProfiles));
                }

            }

            //profile the workbook doesn't know about
            if (segment.IsWorkersComp && segment.WorkersCompStateAttachmentProfile.SourceId.HasValue)
            {
                var localId = segment.WorkersCompStateAttachmentProfile.SourceId.Value;
                var serverOrphan = serverProfiles.SingleOrDefault(sp => sp.Key != localId).Value;

                if (serverOrphan != null) segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.DistributionName));
            }
            else
            {
                if (serverProfiles.Any())
                {
                    segmentView.ChildViews.Add(CreateServerOnlyView(serverProfiles.SingleOrDefault().Value?.DistributionName));
                }
            }
        }

        private static SynchronizationView CreateWorkersCompStateAttachmentProfileView(SegmentModel segmentModel, WorkersCompStateAttachmentProfile profile,
            Dictionary<long, CollectorApi.WorkersCompStateAttachmentDistributionModel> serverProfile)
        {
            var model = segmentModel.WorkersCompStateAttachmentModels.Single(ctm => ctm.Guid == profile.Guid);
            var code = GetWorkersCompStateAttachmentSynchronizationCode(model, serverProfile);

            var profileView = new SynchronizationView
            {
                Name = profile.CommonExcelMatrix.FullName,
                SynchronizationCode = code
            };

            return profileView;
        }


        private void CreateWorkersCompStateHazard(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //profiles the workbook knows about
            var serverProfiles = _bexCommunicationManager.GetServerWorkersCompStateHazardGroupProfiles(_package.SourceId, segment.SourceId);
            if (segment.IsWorkersComp)
            {
                var showProfile = segment.WorkersCompStateHazardGroupProfile.SourceId.HasValue ||
                    segment.WorkersCompStateHazardGroupProfile.ExcelMatrix.Items.Any();
                if (showProfile)
                {
                    segmentView.ChildViews.Add(CreateWorkersCompStateHazardProfileView(segmentModel,
                        segment.WorkersCompStateHazardGroupProfile, serverProfiles));
                }
            }

            //profile the workbook doesn't know about
            if (segment.IsWorkersComp && segment.WorkersCompStateHazardGroupProfile.SourceId.HasValue)
            {
                var localId = segment.WorkersCompStateHazardGroupProfile.SourceId.Value;
                var serverOrphan = serverProfiles.SingleOrDefault(sp => sp.Key != localId).Value;

                if (serverOrphan != null) segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.DistributionName));
            }
            else
            {
                if (serverProfiles.Any())
                {
                    segmentView.ChildViews.Add(CreateServerOnlyView(serverProfiles.SingleOrDefault().Value?.DistributionName));
                }
            }
        }

        private static SynchronizationView CreateWorkersCompStateHazardProfileView(SegmentModel segmentModel, WorkersCompStateHazardGroupProfile profile,
            Dictionary<long, CollectorApi.WorkersCompStateHazardGroupDistributionModel> serverProfile)
        {
            var model = segmentModel.WorkersCompStateHazardGroupModels.Single(ctm => ctm.Guid == profile.Guid);
            var code = GetWorkersCompStateHazardSynchronizationCode(model, serverProfile);

            var profileView = new SynchronizationView
            {
                Name = profile.CommonExcelMatrix.FullName,
                SynchronizationCode = code
            };

            return profileView;
        }


        private void CreateWorkersCompClassCode(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //profiles the workbook knows about
            var serverProfiles = _bexCommunicationManager.GetServerWorkersCompClassCodeProfiles(_package.SourceId, segment.SourceId);
            if (segment.IsWorkersComp && segment.IsWorkerCompClassCodeActive)
            {
                var showProfile = segment.WorkersCompClassCodeProfile.SourceId.HasValue 
                                  || segment.WorkersCompClassCodeProfile.ExcelMatrix.Items.Any();
                if (showProfile)
                {
                    segmentView.ChildViews.Add(CreateWorkersCompClassCodeProfileView(segmentModel,
                        segment.WorkersCompClassCodeProfile, serverProfiles));
                }
            }

            //profile the workbook doesn't know about
            if (segment.IsWorkersComp
                    && segment.IsWorkerCompClassCodeActive 
                    && segment.WorkersCompClassCodeProfile.SourceId.HasValue)
            {
                var localId = segment.WorkersCompClassCodeProfile.SourceId.Value;
                var serverOrphan = serverProfiles.SingleOrDefault(sp => sp.Key != localId).Value;

                if (serverOrphan != null) segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.DistributionName));
            }
            else
            {
                if (serverProfiles.Any())
                {
                    segmentView.ChildViews.Add(CreateServerOnlyView(serverProfiles.SingleOrDefault().Value?.DistributionName));
                }
            }
        }

        private static SynchronizationView CreateWorkersCompClassCodeProfileView(SegmentModel segmentModel, WorkersCompClassCodeProfile profile,
            Dictionary<long, CollectorApi.WorkersCompStateClassCodeDistributionModel> serverProfile)
        {
            var model = segmentModel.WorkersCompClassCodeModels.Single(ctm => ctm.Guid == profile.Guid);
            var code = GetWorkersCompClassCodeSynchronizationCode(model, serverProfile);

            var profileView = new SynchronizationView
            {
                Name = profile.CommonExcelMatrix.FullName,
                SynchronizationCode = code
            };

            return profileView;
        }



        private static SynchronizationCode GetStateSynchronizationCode(StateModel stateModel,
            Dictionary<long, CollectorApi.StateDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!stateModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(stateModel.SourceId.HasValue, "State Id is null");
                var sourceId = stateModel.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorStateModel = stateModel.Map();
                    synchronizationCode = collectorStateModel.IsEqualTo(serverProfiles[sourceId]) ? SynchronizationCode.InSynchronization :
                        stateModel.Items.Any() ? SynchronizationCode.NotInSynchronization : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static SynchronizationCode GetConstructionTypeSynchronizationCode(ConstructionTypeModel constructionTypeModel,
            Dictionary<long, CollectorApi.ConstructionTypeDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!constructionTypeModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(constructionTypeModel.SourceId.HasValue, "Construction Type Id is null");
                var sourceId = constructionTypeModel.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorConstructionTypeModel = constructionTypeModel.Map();
                    synchronizationCode = collectorConstructionTypeModel.IsEqualTo(serverProfiles[sourceId])
                        ?
                        SynchronizationCode.InSynchronization
                        : constructionTypeModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static SynchronizationCode GetOccupancyTypeSynchronizationCode(OccupancyTypeModel occupancyTypeModel,
            Dictionary<long, CollectorApi.OccupancyTypeDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!occupancyTypeModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(occupancyTypeModel.SourceId.HasValue, "Occupancy Type Id is null");
                var sourceId = occupancyTypeModel.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorOccupancyTypeModel = occupancyTypeModel.Map();
                    synchronizationCode = collectorOccupancyTypeModel.IsEqualTo(serverProfiles[sourceId])
                        ?
                        SynchronizationCode.InSynchronization
                        : occupancyTypeModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static SynchronizationCode GetProtectionClassSynchronizationCode(ProtectionClassModel protectionClassModel,
            Dictionary<long, CollectorApi.ProtectionClassDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!protectionClassModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(protectionClassModel.SourceId.HasValue, "Protection Class Id is null");
                var sourceId = protectionClassModel.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorProtectionClassModel = protectionClassModel.Map();
                    synchronizationCode = collectorProtectionClassModel.IsEqualTo(serverProfiles[sourceId])
                        ?
                        SynchronizationCode.InSynchronization
                        : protectionClassModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static SynchronizationCode GetTotalInsuredValueSynchronizationCode(TotalInsuredValueModel totalInsuredValueModel,
            Dictionary<long, CollectorApi.TotalInsuredValueDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!totalInsuredValueModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(totalInsuredValueModel.SourceId.HasValue, "Total Insured Value Id is null");
                var sourceId = totalInsuredValueModel.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorTotalInsuredValueModel = totalInsuredValueModel.Map();
                    synchronizationCode = collectorTotalInsuredValueModel.IsEqualTo(serverProfiles[sourceId])
                        ?
                        SynchronizationCode.InSynchronization
                        : totalInsuredValueModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }


        
        private static SynchronizationCode GetWorkersCompStateAttachmentSynchronizationCode(WorkersCompStateAttachmentModel model,
            Dictionary<long, CollectorApi.WorkersCompStateAttachmentDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!model.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(model.SourceId.HasValue, "WC State Attachment Id is null");
                var sourceId = model.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorWorkersCompStateAttachmentModel = model.Map();
                    synchronizationCode = collectorWorkersCompStateAttachmentModel.IsEqualTo(serverProfiles[sourceId])
                        ?
                        SynchronizationCode.InSynchronization
                        : model.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static SynchronizationCode GetWorkersCompStateHazardSynchronizationCode(WorkersCompStateHazardGroupModel groupModel,
            Dictionary<long, CollectorApi.WorkersCompStateHazardGroupDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!groupModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(groupModel.SourceId.HasValue, "WC State Hazard Id is null");
                var sourceId = groupModel.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorModel = groupModel.Map();
                    synchronizationCode = collectorModel.IsEqualTo(serverProfiles[sourceId])
                        ?
                        SynchronizationCode.InSynchronization
                        : groupModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static SynchronizationCode GetWorkersCompClassCodeSynchronizationCode(WorkersCompClassCodeModel model,
            Dictionary<long, CollectorApi.WorkersCompStateClassCodeDistributionModel> serverProfiles)
        {
            SynchronizationCode synchronizationCode;
            if (!model.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(model.SourceId.HasValue, "WC State Class Code is null");
                var sourceId = model.SourceId.Value;
                if (serverProfiles.ContainsKey(sourceId))
                {
                    var collectorModel = model.Map();
                    synchronizationCode = collectorModel.IsEqualTo(serverProfiles[sourceId])
                        ?
                        SynchronizationCode.InSynchronization
                        : model.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }


        private void CreateExposureSet(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //exposureSet the workbook knows about
            var serverExposureSets = _bexCommunicationManager.GetServerExposureSets(_package.SourceId, segment.SourceId);
            var setsToHide = segment.ExposureSets.Where(es => !es.SourceId.HasValue && !es.ExcelMatrix.Items.Any());
            segment.ExposureSets.OrderBy(es => es.IntraDisplayOrder)
                .Except(setsToHide)
                .ForEach(es =>
                {
                    var exposureSetModel = segmentModel.ExposureSetModels.Single(esm => esm.Guid == es.Guid);
                    var view = CreateExposureSetView(exposureSetModel, serverExposureSets);

                    //test all of the exposures
                    if (view.SynchronizationCode == SynchronizationCode.InSynchronization)
                    {
                        var serverExposures = _bexCommunicationManager.GetServerExposures(_package.SourceId, segment.SourceId, es.SourceId);
                        Debug.Assert(es.SourceId != null, "Exposure set id is null");
                        if (!AreExposuresSynchronized(exposureSetModel.Items, serverExposures))
                        {
                            view.SynchronizationCode = SynchronizationCode.NotInSynchronization;
                        }
                    }

                    segmentView.ChildViews.Add(view);
                });

            //exposureSet the workbook doesn't know about
            var localIds = segment.ExposureSets.Where(es => es.SourceId.HasValue).Select(es => es.SourceId.Value);
            var serverOrphans = serverExposureSets.Where(serverProfile => !localIds.Contains(serverProfile.Key));
            serverOrphans.ForEach(serverOrphan => segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.Name)));
        }

        private static SynchronizationView CreateExposureSetView(ExposureSetModel exposureSetModel,
            Dictionary<long, CollectorApi.ExposureSetModel> serverExposureSets)
        {
            var synchronizationCode = GetExposureSetSynchronizationCode(exposureSetModel, serverExposureSets);

            var exposureSetView = new SynchronizationView {Name = exposureSetModel.Name, SynchronizationCode = synchronizationCode};

            return exposureSetView;
        }

        private static SynchronizationCode GetExposureSetSynchronizationCode(ExposureSetModel exposureSetModel,
            Dictionary<long, CollectorApi.ExposureSetModel> serverExposureSets)
        {
            SynchronizationCode synchronizationCode;
            if (!exposureSetModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(exposureSetModel.SourceId.HasValue, "Exposure Set Id is null");
                var exposureSetId = exposureSetModel.SourceId.Value;
                if (serverExposureSets.ContainsKey(exposureSetId))
                {
                    var collectorExposureSetModel = exposureSetModel.Map();
                    synchronizationCode = collectorExposureSetModel.IsEqualTo(serverExposureSets[exposureSetId])
                        ?
                        SynchronizationCode.InSynchronization
                        : exposureSetModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static bool AreExposuresSynchronized(IList<ExposureModelPlus> exposures,
            Dictionary<long, CollectorApi.ExposureModel> serverExposures)
        {
            if (exposures.Count != serverExposures.Count) return false;
            foreach (var exposure in exposures)
            {
                if (!exposure.SourceId.HasValue) return false;
                if (!serverExposures.ContainsKey(exposure.SourceId.Value)) return false;
                if (!exposure.IsEqualTo(serverExposures[exposure.SourceId.Value])) return false;
            }

            return true;
        }

        private void CreateAggregateLossSet(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //aggregateLossSet the workbook knows about
            var serverAggregateLossSets = _bexCommunicationManager.GetServerAggregateLossSets(_package.SourceId, segment.SourceId);
            var setsToHide = segment.AggregateLossSets.Where(es => !es.SourceId.HasValue && !es.ExcelMatrix.Items.Any());
            foreach (var als in segment.AggregateLossSets.OrderBy(als => als.IntraDisplayOrder).Except(setsToHide))
            {
                var aggregateLossSetModel = segmentModel.AggregateLossSetModels.Single(model => model.Guid == als.Guid);
                var view = CreateAggregateLossSetView(segmentModel, als, serverAggregateLossSets);

                //test all of the aggregate losses
                if (view.SynchronizationCode == SynchronizationCode.InSynchronization)
                {
                    var serverAggregateLosses =
                        _bexCommunicationManager.GetServerAggregateLosses(_package.SourceId, segment.SourceId, als.SourceId);
                    Debug.Assert(als.SourceId != null, "Aggregate loss set id is null");
                    if (!AreAggregateLossesSynchronized(aggregateLossSetModel.Items, serverAggregateLosses))
                    {
                        view.SynchronizationCode = SynchronizationCode.NotInSynchronization;
                    }
                }

                segmentView.ChildViews.Add(view);
            }

            //aggregateLossSet the workbook doesn't know about
            var localIds = segment.AggregateLossSets.Where(es => es.SourceId.HasValue).Select(es => es.SourceId.Value);
            var serverOrphans = serverAggregateLossSets.Where(serverLossSet => !localIds.Contains(serverLossSet.Key));
            serverOrphans.ForEach(serverOrphan => segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.Name)));
        }

        private static SynchronizationView CreateAggregateLossSetView(SegmentModel segmentModel, AggregateLossSet aggregateLossSetProfile,
            Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.AggregateLossSetModel> serverLossSets)
        {
            var aggregateLossSetModel = segmentModel.AggregateLossSetModels.Single(alm => alm.Guid == aggregateLossSetProfile.Guid);
            var synchronizationCode = GetAggregateLossSetSynchronizationCode(aggregateLossSetModel, serverLossSets);

            var aggregateLossSetView = new SynchronizationView
            {
                Name = aggregateLossSetProfile.CommonExcelMatrix.FullName, SynchronizationCode = synchronizationCode
            };

            return aggregateLossSetView;
        }

        private static bool AreAggregateLossesSynchronized(List<AggregateLossModelPlus> aggregateLosses,
            Dictionary<long, CollectorApi.AggregateLossModel> serverAggregateLosses)
        {
            if (aggregateLosses.Count != serverAggregateLosses.Count) return false;
            foreach (var loss in aggregateLosses)
            {
                if (!loss.SourceId.HasValue) return false;
                if (!serverAggregateLosses.ContainsKey(loss.SourceId.Value)) return false;
                if (!loss.IsEqualTo(serverAggregateLosses[loss.SourceId.Value])) return false;
            }

            return true;
        }

        private static SynchronizationCode GetAggregateLossSetSynchronizationCode(AggregateLossSetModel aggregateLossSetModel,
            Dictionary<long, CollectorApi.AggregateLossSetModel> serverLossSets)
        {
            SynchronizationCode synchronizationCode;
            if (!aggregateLossSetModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(aggregateLossSetModel.SourceId.HasValue, "Aggregate Loss Set Id is null");
                var aggregateLossSetId = aggregateLossSetModel.SourceId.Value;
                if (serverLossSets.ContainsKey(aggregateLossSetId))
                {
                    var collectorAggregateLossSetModel = aggregateLossSetModel.Map();
                    synchronizationCode = collectorAggregateLossSetModel.IsEqualTo(serverLossSets[aggregateLossSetId])
                        ?
                        SynchronizationCode.InSynchronization
                        : aggregateLossSetModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private void CreateIndividualLossSet(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //individualLossSet the workbook knows about
            var serverIndividualLossSets = _bexCommunicationManager.GetServerIndividualLossSets(_package.SourceId, segment.SourceId);
            var setsToHide = segment.IndividualLossSets.Where(es => !es.SourceId.HasValue && !es.ExcelMatrix.Items.Any());

            foreach (var ils in segment.IndividualLossSets.OrderBy(ils => ils.IntraDisplayOrder).Except(setsToHide))
            {
                var individualLossSetModel = segmentModel.IndividualLossSetModels.Single(model => model.Guid == ils.Guid);
                var view = CreateIndividualLossSetView(segmentModel, ils, serverIndividualLossSets);

                //test all of the individual losses
                if (view.SynchronizationCode == SynchronizationCode.InSynchronization)
                {
                    var serverIndividualLosses =
                        _bexCommunicationManager.GetServerIndividualLosses(_package.SourceId, segment.SourceId, ils.SourceId);
                    Debug.Assert(ils.SourceId != null, "Individual loss set id is null");
                    if (!AreIndividualLossesSynchronized(individualLossSetModel.Items, serverIndividualLosses))
                    {
                        view.SynchronizationCode = SynchronizationCode.NotInSynchronization;
                    }
                }

                segmentView.ChildViews.Add(view);
            }

            //individualLossSet the workbook doesn't know about
            var localIds = segment.IndividualLossSets.Where(es => es.SourceId.HasValue).Select(es => es.SourceId.Value);
            var serverOrphans = serverIndividualLossSets.Where(serverLossSet => !localIds.Contains(serverLossSet.Key));
            serverOrphans.ForEach(serverOrphan => segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.Name)));
        }

        private static SynchronizationView CreateIndividualLossSetView(SegmentModel segmentModel,
            IndividualLossSet individualLossSetProfile, Dictionary<long, CollectorApi.IndividualLossSetModel> serverLossSets)
        {
            var individualLossSetModel = segmentModel.IndividualLossSetModels.Single(ilm => ilm.Guid == individualLossSetProfile.Guid);
            var synchronizationCode = GetIndividualLossSetSynchronizationCode(individualLossSetModel, serverLossSets);

            var individualLossSetView = new SynchronizationView
            {
                Name = individualLossSetProfile.CommonExcelMatrix.FullName, SynchronizationCode = synchronizationCode
            };

            return individualLossSetView;
        }

        private static bool AreIndividualLossesSynchronized(List<IndividualLossModelPlus> individualLosses,
            Dictionary<long, CollectorApi.IndividualLossModel> serverIndividualLosses)
        {
            if (individualLosses.Count != serverIndividualLosses.Count) return false;
            foreach (var loss in individualLosses)
            {
                if (!loss.SourceId.HasValue) return false;
                if (!serverIndividualLosses.ContainsKey(loss.SourceId.Value)) return false;
                if (!loss.IsEqualTo(serverIndividualLosses[loss.SourceId.Value])) return false;
            }

            return true;
        }

        private static SynchronizationCode GetIndividualLossSetSynchronizationCode(IndividualLossSetModel individualLossSetModel,
            Dictionary<long, CollectorApi.IndividualLossSetModel> serverLossSets)
        {
            SynchronizationCode synchronizationCode;
            if (!individualLossSetModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(individualLossSetModel.SourceId.HasValue, "Individual Loss Set Id is null");
                var individualLossSetId = individualLossSetModel.SourceId.Value;
                if (serverLossSets.ContainsKey(individualLossSetId))
                {
                    var collectorIndividualLossSetModel = individualLossSetModel.Map();
                    synchronizationCode = collectorIndividualLossSetModel.IsEqualTo(serverLossSets[individualLossSetId])
                        ?
                        SynchronizationCode.InSynchronization
                        : individualLossSetModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private void CreateRateChangeSet(ISegment segment, SegmentModel segmentModel, SynchronizationView segmentView)
        {
            //rate change set the workbook knows about
            var serverRateChangeSets = _bexCommunicationManager.GetServerRateChangeSets(_package.SourceId, segment.SourceId);
            var setsToHide = segment.RateChangeSets.Where(es => !es.SourceId.HasValue && !es.ExcelMatrix.Items.Any());

            foreach (var rateChangeSet in segment.RateChangeSets.OrderBy(ils => ils.IntraDisplayOrder).Except(setsToHide))
            {
                var rateChangeSetModel = segmentModel.RateChangeSetModels.Single(model => model.Guid == rateChangeSet.Guid);
                var view = CreateRateChangeSetView(segmentModel, rateChangeSet, serverRateChangeSets);

                //test all of the rate changes
                if (view.SynchronizationCode == SynchronizationCode.InSynchronization)
                {
                    var serverRateChanges =
                        _bexCommunicationManager.GetServerRateChanges(_package.SourceId, segment.SourceId, rateChangeSet.SourceId);
                    Debug.Assert(rateChangeSet.SourceId != null, "Rate change set id is null");
                    if (!AreRateChangesSynchronized(rateChangeSetModel.Items, serverRateChanges))
                    {
                        view.SynchronizationCode = SynchronizationCode.NotInSynchronization;
                    }
                }

                segmentView.ChildViews.Add(view);
            }

            //rate change set the workbook doesn't know about
            var localIds = segment.RateChangeSets.Where(es => es.SourceId.HasValue).Select(es => es.SourceId.Value);
            var serverOrphans = serverRateChangeSets.Where(set => !localIds.Contains(set.Key));
            serverOrphans.ForEach(serverOrphan => segmentView.ChildViews.Add(CreateServerOnlyView(serverOrphan.Value.Name)));
        }

        private static SynchronizationView CreateRateChangeSetView(SegmentModel segmentModel, RateChangeSet rateChangeSet,
            Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.RateChangeSetModel> serverLossSets)
        {
            var rateChangeSetModel = segmentModel.RateChangeSetModels.Single(model => model.Guid == rateChangeSet.Guid);
            var synchronizationCode = GetRateChangeSetSynchronizationCode(rateChangeSetModel, serverLossSets);

            var rateChangeSetView = new SynchronizationView
            {
                Name = rateChangeSet.CommonExcelMatrix.FullName, SynchronizationCode = synchronizationCode
            };

            return rateChangeSetView;
        }

        private static bool AreRateChangesSynchronized(List<RateChangeModelPlus> rateChanges,
            Dictionary<long, CollectorApi.RateChangeModel> serverRateChanges)
        {
            if (rateChanges.Count != serverRateChanges.Count) return false;
            foreach (var rateChange in rateChanges)
            {
                if (!rateChange.SourceId.HasValue) return false;
                if (!serverRateChanges.ContainsKey(rateChange.SourceId.Value)) return false;
                if (!rateChange.IsEqualTo(serverRateChanges[rateChange.SourceId.Value])) return false;
            }

            return true;
        }

        private static SynchronizationCode GetRateChangeSetSynchronizationCode(RateChangeSetModel rateChangeSetModel,
            Dictionary<long, CollectorApi.RateChangeSetModel> serverRateChangeSets)
        {
            SynchronizationCode synchronizationCode;
            if (!rateChangeSetModel.SourceId.HasValue)
            {
                synchronizationCode = SynchronizationCode.New;
            }
            else
            {
                Debug.Assert(rateChangeSetModel.SourceId.HasValue, "Rate change set Id is null");
                var rateChangeSetId = rateChangeSetModel.SourceId.Value;
                if (serverRateChangeSets.ContainsKey(rateChangeSetId))
                {
                    var collectorRateChangeSetModel = rateChangeSetModel.Map();
                    synchronizationCode = collectorRateChangeSetModel.IsEqualTo(serverRateChangeSets[rateChangeSetId])
                        ?
                        SynchronizationCode.InSynchronization
                        : rateChangeSetModel.Items.Any()
                            ? SynchronizationCode.NotInSynchronization
                            : SynchronizationCode.Deleted;
                }
                else
                {
                    synchronizationCode = SynchronizationCode.New;
                }
            }

            return synchronizationCode;
        }

        private static SynchronizationView CreateServerOnlyView(string name)
        {
            return new SynchronizationView {Name = name, SynchronizationCode = SynchronizationCode.Deleted};
        }
    }

    public abstract class BaseSynchronizationViewModel : ViewModelBase
    {
        private double _treeFontSize;
        public double DefaultTreeFontSize = 14;
        private GridLength _validationRowPixels;
        private GridLength _loadingRowPixels;
        private GridLength _bodyRowPixels;

        protected GridLength HeadingRowHeight = new GridLength(40);
        protected GridLength LoadingRowHeight = new GridLength(80);
        protected GridLength ValidationRowHeight = new GridLength(80);
        protected GridLength ErrorRowHeight = new GridLength(80);
        protected GridLength BodyRowHeight = new GridLength(1, GridUnitType.Star);
        protected GridLength LegendRowHeight = new GridLength(125);
        protected GridLength BottomRowHeight = new GridLength(40);

        private ObservableCollection<SynchronizationView> _packageSynchronizationViews;
        private GridLength _legendRowPixels;
        private string _errorMessage;
        private string _statusMessage;
        private bool _showZoom;
        private GridLength _headingRowPixels;
        private GridLength _bottomRowPixels;
        private string _validationMessage;
        private GridLength _errorRowPixels;
        private string _title;

        public virtual bool IsEntirePackageSynchronized => true;

        public ObservableCollection<SynchronizationView> PackageSynchronizationViews
        {
            get => _packageSynchronizationViews;
            set
            {
                _packageSynchronizationViews = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength ValidationRowPixels
        {
            get => _validationRowPixels;
            set
            {
                _validationRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength ErrorRowPixels
        {
            get => _errorRowPixels;
            set
            {
                _errorRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength LoadingRowPixels
        {
            get => _loadingRowPixels;
            set
            {
                _loadingRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength HeadingRowPixels
        {
            get => _headingRowPixels;
            set
            {
                _headingRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength BodyRowPixels
        {
            get => _bodyRowPixels;
            set
            {
                _bodyRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength LegendRowPixels
        {
            get => _legendRowPixels;
            set
            {
                _legendRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength BottomRowPixels
        {
            get => _bottomRowPixels;
            set
            {
                _bottomRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public string Title
        {
            get => _title;
            set
            {
                _title = value;
                NotifyPropertyChanged();
            }
        }

        public string ValidationMessage
        {
            get => _validationMessage;
            set
            {
                _validationMessage = value;
                NotifyPropertyChanged();
            }
        }

        public string ErrorMessage
        {
            get => _errorMessage;
            set
            {
                _errorMessage = value;
                NotifyPropertyChanged();
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                _statusMessage = value;
                NotifyPropertyChanged();
            }
        }

        public bool ShowZoom
        {
            get => _showZoom;
            set
            {
                _showZoom = value;
                NotifyPropertyChanged();
            }
        }

        public double TreeFontSize
        {
            get => _treeFontSize;
            set
            {
                _treeFontSize = value;
                NotifyPropertyChanged();
            }
        }

        protected BaseSynchronizationViewModel()
        {
            PackageSynchronizationViews = new ObservableCollection<SynchronizationView>();
            TreeFontSize = DefaultTreeFontSize;
            Title = BexConstants.UploadPreviewName;
        }
    }

    public class SynchronizationView : ViewModelBase
    {
        public string Name { get; set; }
        public IList<SynchronizationView> ChildViews { get; set; }
        public SynchronizationCode SynchronizationCode { get; set; }

        public SynchronizationView()
        {
            ChildViews = new List<SynchronizationView>();
        }
    }

    public enum SynchronizationCode
    {
        New,
        Deleted,
        InSynchronization,
        NotInSynchronization,
        None
    }
}