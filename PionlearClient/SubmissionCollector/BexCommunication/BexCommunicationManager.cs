using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using MunichRe.Bex.ApiClient;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient;
using PionlearClient.CollectorClientPlus;
using PionlearClient.CollectorClientPlus.Extensions;
using PionlearClient.Extensions;
using PionlearClient.Model;
using PionlearClient.TokenAuthentication;
using SubmissionCollector.ExcelWorkspaceFolder;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Segment;
using AggregateLossSetModel = PionlearClient.Model.AggregateLossSetModel;
using ExposureSetModel = PionlearClient.Model.ExposureSetModel;
using IndividualLossSetModel = PionlearClient.Model.IndividualLossSetModel;
using RateChangeSetModel = PionlearClient.Model.RateChangeSetModel;

// ReSharper disable PossibleInvalidOperationException

namespace SubmissionCollector.BexCommunication
{
    public class BexCommunicationManager
    {
        private readonly CollectorClient _collectorClient;
        private ActivityTracker _activityTracker;

        public BexCommunicationManager()
        {
            _collectorClient = BexCollectorClientFactory.CreateBexCollectorClient(ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
        }

        internal IEnumerable<SubmissionPackageAttachedTo> GetRatingAnalysisIdsUsingThisPackage(long sourceId)
        {
            var task = _collectorClient.SubmissionPackagesClient.AttachedStatusAsync(sourceId);
            try
            {
                task.Wait();
            }
            catch (AggregateException ex)
            {
                var swaggerException = ex.InnerException as SwaggerException;
                var notFoundAsInteger = Convert.ToInt32(HttpStatusCode.NotFound);
                if (swaggerException != null && swaggerException.StatusCode == notFoundAsInteger)
                {
                    throw new FileNotFoundException("Can't find URL.");
                }

                throw;
            }

            return task.Result;
        }

        internal void UploadModels(PackageModel packageModel, ActivityTracker activityTracker, BackgroundWorker worker,
            StackTraceLogger logger)
        {
            try
            {
                UploadModels(packageModel, activityTracker, worker);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                throw;
            }
        }

        private void UploadModels(PackageModel packageModel, ActivityTracker activityTracker, BackgroundWorker worker)
        {
            _activityTracker = activityTracker;

            worker?.ReportProgress(0, $"Uploading {BexConstants.PackageName.ToLower()} <{packageModel.Name}>...");
            InsertAndUpdatePackage(packageModel);

            InsertAndUpdateSegments(packageModel);
            DeleteSegments(packageModel);

            foreach (var segmentModel in packageModel.SegmentModels)
            {
                worker?.ReportProgress(0, $"Uploading {BexConstants.SegmentName.ToLower()} <{segmentModel.Name}> ...");

                var packageModelSourceId = packageModel.SourceId.Value;

                DeleteHazardProfiles(packageModelSourceId, segmentModel);
                InsertAndUpdateHazardProfiles(packageModelSourceId, segmentModel);

                DeletePolicyProfiles(packageModelSourceId, segmentModel);
                InsertAndUpdatePolicyProfiles(packageModelSourceId, segmentModel);

                DeleteStateProfiles(packageModelSourceId, segmentModel);
                InsertAndUpdateStateProfiles(packageModelSourceId, segmentModel);

                
                DeleteProtectionClassProfiles(packageModelSourceId, segmentModel);
                InsertAndUpdateProtectionClassProfiles(packageModelSourceId, segmentModel);

                DeleteConstructionTypeProfiles(packageModelSourceId, segmentModel);
                InsertAndUpdateConstructionTypeProfiles(packageModelSourceId, segmentModel); 

                DeleteOccupancyTypeProfiles(packageModelSourceId, segmentModel);
                InsertAndUpdateOccupancyTypeProfiles(packageModelSourceId, segmentModel);

                DeleteTotalInsuredValueProfiles(packageModelSourceId, segmentModel);
                InsertAndUpdateTotalInsuredValueProfiles(packageModelSourceId, segmentModel);



                DeleteWorkersCompStateAttachmentProfile(packageModelSourceId, segmentModel);
                InsertAndUpdateWorkersCompStateAttachmentProfile(packageModelSourceId, segmentModel);

                DeleteWorkersCompStateHazardGroupProfile(packageModelSourceId, segmentModel);
                InsertAndUpdateWorkersCompStateHazardGroupProfile(packageModelSourceId, segmentModel);

                DeleteWorkersCompClassCodeProfile(packageModelSourceId, segmentModel);
                InsertAndUpdateWorkersCompClassCodeProfile(packageModelSourceId, segmentModel);



                DeleteExposureSets(packageModelSourceId, segmentModel);
                InsertAndUpdateExposureSets(packageModelSourceId, segmentModel);
                foreach (var exposureSetModel in segmentModel.ExposureSetModels.Where(esm => esm.Items.Any()))
                {
                    DeleteExposures(packageModelSourceId, segmentModel, exposureSetModel);
                    InsertAndUpdateExposures(packageModelSourceId, segmentModel, exposureSetModel);
                }

                DeleteAggregateLossSets(packageModelSourceId, segmentModel);
                InsertAndUpdateAggregateLossSets(packageModelSourceId, segmentModel);
                foreach (var aggregateLossSetModel in segmentModel.AggregateLossSetModels.Where(alm => alm.Items.Any()))
                {
                    DeleteAggregateLosses(packageModelSourceId, segmentModel, aggregateLossSetModel);
                    InsertAndUpdateAggregateLosses(packageModelSourceId, segmentModel, aggregateLossSetModel);
                }

                var sb = new StringBuilder();
                sb.AppendLine($"Uploading {BexConstants.SegmentName.ToLower()} <{segmentModel.Name}>");
                sb.AppendLine($"{BexConstants.IndividualLossSetName.ToLower()} ...");
                worker?.ReportProgress(0, sb.ToString());

                DeleteIndividualLossSets(packageModelSourceId, segmentModel);
                InsertAndUpdateIndividualLossSets(packageModelSourceId, segmentModel);
                foreach (var individualLossSetModel in segmentModel.IndividualLossSetModels.Where(ilm => ilm.Items.Any()))
                {
                    DeleteIndividualLosses(packageModelSourceId, segmentModel, individualLossSetModel);
                    InsertAndUpdateIndividualLosses(packageModelSourceId, segmentModel, individualLossSetModel, worker);
                }

                sb.Clear();
                sb.AppendLine($"Uploading {BexConstants.SegmentName.ToLower()} <{segmentModel.Name}>");
                sb.AppendLine($"{BexConstants.RateChangeSetName.ToLower()} ...");
                worker?.ReportProgress(0, sb.ToString());

                DeleteRateChangeSets(packageModelSourceId, segmentModel);
                InsertAndUpdateRateChangeSets(packageModelSourceId, segmentModel);
                foreach (var rateChangeSetModel in segmentModel.RateChangeSetModels.Where(ilm => ilm.Items.Any()))
                {
                    DeleteRateChanges(packageModelSourceId, segmentModel, rateChangeSetModel);
                    InsertAndUpdateRateChanges(packageModelSourceId, segmentModel, rateChangeSetModel, worker);
                }
            }
        }

        private void InsertAndUpdatePackage(PackageModel packageModel)
        {
            var collectorPackageModel = packageModel.Map();
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;

            if (!packageModel.SourceId.HasValue)
            {
                InsertPackage(packageModel, collectorPackageModel, package);
            }
            else
            {
                var packageModelSourceId = packageModel.SourceId.Value;
                var serverPackageModel = TryToGetSubmissionPackageModel(packageModelSourceId);
                if (serverPackageModel == null)
                {
                    collectorPackageModel.Id = 0;
                    InsertPackage(packageModel, collectorPackageModel, package);
                    return;
                }

                if (!collectorPackageModel.IsEqualTo(serverPackageModel))
                {
                    UpdatePackage(packageModel, collectorPackageModel, package);
                }
            }
        }

        public SubmissionPackageModel TryToGetSubmissionPackageModel(long packageModelSourceId)
        {
            try
            {
                return GetSubmissionPackageModel(packageModelSourceId);
            }
            catch (AggregateException ae)
            {
                if (!ae.InnerExceptions.First().Message.Contains(BexConstants.NotFoundErrorCode)) throw;
                return null;
            }
        }


        private SubmissionPackageModel GetSubmissionPackageModel(long packageModelSourceId)
        {
            var client = _collectorClient.SubmissionPackagesClient;
            client.Logger = new ClientLogger();

            var getTask = client.GetAsync(packageModelSourceId);
            getTask.Wait();
            return getTask.Result;
        }

        private void UpdatePackage(PackageModel packageModel, SubmissionPackageModel collectorPackageModel,
            Models.Package.Package package)
        {
            var client = _collectorClient.SubmissionPackagesClient;

            client.Logger = new ClientLogger();
            var putTask = client.PutAsync(collectorPackageModel, collectorPackageModel.Id.Value);
            putTask.Wait();

            packageModel.SourceTimestamp = package.SourceTimestamp = DateTime.Now;
            packageModel.IsDirty = package.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 0, Name = packageModel.Name});
        }

        private void InsertPackage(PackageModel packageModel, SubmissionPackageModel collectorPackageModel,
            Models.Package.Package package)
        {
            var postTask = _collectorClient.SubmissionPackagesClient.PostAsync(collectorPackageModel);
            postTask.Wait();

            packageModel.SourceId = package.SourceId = postTask.Result;
            packageModel.SourceTimestamp = package.SourceTimestamp = DateTime.Now;
            packageModel.IsDirty = package.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Insert, Priority = 0, Name = packageModel.Name});
        }

        private void DeleteSegments(PackageModel packageModel)
        {
            var getAllTask = _collectorClient.SubmissionSegmentsClient.GetAllAsync(packageModel.SourceId.Value);
            getAllTask.Wait();

            var serverDictionary = getAllTask.Result.ToDictionary(k => k.Id.Value, v => v.Name);
            var localSegmentIds = packageModel.SegmentModels.Select(sm => sm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localSegmentIds))
            {
                var deleteTask = _collectorClient.SubmissionSegmentsClient.DeleteAsync(packageModel.SourceId.Value, id);
                deleteTask.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 1, Name = name, SegmentId = id});
            }
        }

        private void InsertAndUpdateSegments(PackageModel packageModel)
        {
            var packageModelSourceId = packageModel.SourceId.Value;
            var serverSegmentDictionary = GetServerSegments(packageModelSourceId);

            foreach (var segmentModel in packageModel.SegmentModels)
            {
                var collectorSegmentModel = segmentModel.Map();

                if (!segmentModel.SourceId.HasValue)
                {
                    InsertSegment(packageModelSourceId, segmentModel, collectorSegmentModel);
                }
                else
                {
                    var segmentId = collectorSegmentModel.Id.Value;
                    if (!serverSegmentDictionary.ContainsKey(segmentId))
                    {
                        collectorSegmentModel.Id = 0;
                        InsertSegment(packageModelSourceId, segmentModel, collectorSegmentModel);
                        continue;
                    }

                    if (collectorSegmentModel.IsEqualTo(serverSegmentDictionary[segmentId])) continue;
                    UpdateSegment(packageModelSourceId, segmentModel, collectorSegmentModel, segmentId);
                }
            }
        }

        public Dictionary<long, SubmissionSegmentModel> GetServerSegments(long? packageModelSourceId)
        {
            if (!packageModelSourceId.HasValue) return new Dictionary<long, SubmissionSegmentModel>();

            var getAllTask = _collectorClient.SubmissionSegmentsClient.GetAllAsync(packageModelSourceId.Value);
            getAllTask.Wait();
            var serverSegments = getAllTask.Result;
            var serverSegmentDictionary = serverSegments.ToDictionary(serverSegment => serverSegment.Id.Value);
            return serverSegmentDictionary;
        }

        private void UpdateSegment(long packageId, SegmentModel segmentModel, SubmissionSegmentModel collectorSegmentModel,
            long segmentId)
        {
            var task = _collectorClient.SubmissionSegmentsClient.PutAsync(packageId, segmentId, collectorSegmentModel);
            task.Wait();

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            segmentModel.SourceTimestamp = segment.SourceTimestamp = DateTime.Now;
            segmentModel.IsDirty = segment.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 1, Name = segmentModel.Name, SegmentId = segmentId});
        }

        private void InsertSegment(long packageId, SegmentModel segmentModel, SubmissionSegmentModel collectorSegmentModel)
        {
            var task = _collectorClient.SubmissionSegmentsClient.PostAsync(packageId, collectorSegmentModel);
            task.Wait();

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            segmentModel.SourceId = segment.SourceId = task.Result;
            segmentModel.SourceTimestamp = segment.SourceTimestamp = DateTime.Now;
            segmentModel.IsDirty = segment.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Insert, Priority = 1, Name = segmentModel.Name, SegmentId = segment.SourceId});
        }

        private void DeleteHazardProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.HazardDistributionClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.HazardModels.Where(hm => hm.SourceId.HasValue).ToList();
            var localHazardModelIds = localModelsWithSourceId.Select(hm => hm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localHazardModelIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segmentId});
            }

            foreach (var hazardModel in localModelsWithSourceId.Where(hazardModel => !hazardModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, hazardModel.SourceId.Value);
                task2.Wait();

                hazardModel.SourceId = null;
                hazardModel.SourceTimestamp = null;
                
                var hazardProfile = segment.HazardProfiles.Single(profile => profile.Guid == hazardModel.Guid);
                hazardProfile.SourceId = null;
                hazardProfile.SourceTimestamp = null;
                
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = hazardModel.Name, SegmentId = segment.SourceId});
            }
        }

        private void InsertAndUpdateHazardProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverHazardDictionary = GetServerHazardProfiles(packageId, segmentId);

            foreach (var hazardModel in segmentModel.HazardModels.Where(hazardModel => hazardModel.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorHazardModel = hazardModel.Map();

                if (!hazardModel.SourceId.HasValue)
                {
                    InsertHazardProfile(packageId, segmentId, hazardModel, segment, collectorHazardModel);
                }
                else
                {
                    var hazardId = hazardModel.SourceId.Value;

                    if (!serverHazardDictionary.ContainsKey(hazardId))
                    {
                        collectorHazardModel.Id = 0;
                        InsertHazardProfile(packageId, segmentId, hazardModel, segment, collectorHazardModel);
                        continue;
                    }

                    if (collectorHazardModel.IsEqualTo(serverHazardDictionary[hazardId])) continue;
                    UpdateHazardProfile(packageId, segmentId, hazardModel, segment, collectorHazardModel, hazardId);
                }
            }
        }

        public Dictionary<long, HazardDistributionModel> GetServerHazardProfiles(long? packageModelSourceId, long? segmentId)
        {
            if (!packageModelSourceId.HasValue || !segmentId.HasValue) return new Dictionary<long, HazardDistributionModel>();

            var getAllTask = _collectorClient.HazardDistributionClient.GetAllAsync(packageModelSourceId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverHazardDistributions = getAllTask.Result;
            var serverHazardDictionary =
                serverHazardDistributions.ToDictionary(hazardDistributionModel => hazardDistributionModel.Id.Value);
            return serverHazardDictionary;
        }

        private void UpdateHazardProfile(long packageModelSourceId, long segmentId, HazardModel hazardModel, ISegment segment,
            HazardDistributionModel collectorHazardModel, long hazardId)
        {
            var task = _collectorClient.HazardDistributionClient.PutAsync(packageModelSourceId, segmentId, hazardId, collectorHazardModel);
            task.Wait();

            var hazardProfile = segment.HazardProfiles.Single(s => s.Guid == hazardModel.Guid);
            hazardModel.SourceTimestamp = hazardProfile.SourceTimestamp = DateTime.Now;
            hazardModel.IsDirty = hazardProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 2, Name = hazardModel.Name, SegmentId = segment.SourceId});
        }

        private void InsertHazardProfile(long packageModelSourceId, long segmentId, HazardModel hazardModel, ISegment segment,
            HazardDistributionModel collectorHazardModel)
        {
            var task = _collectorClient.HazardDistributionClient.PostAsync(packageModelSourceId, segmentId, collectorHazardModel);
            task.Wait();

            var hazardProfile = segment.HazardProfiles.Single(s => s.Guid == hazardModel.Guid);
            hazardModel.SourceId = hazardProfile.SourceId = task.Result;
            hazardModel.SourceTimestamp = hazardProfile.SourceTimestamp = DateTime.Now;
            hazardModel.IsDirty = hazardProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert, Priority = 2, Name = hazardModel.Name, SegmentId = segment.SourceId
            });
        }

        private void DeletePolicyProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.PolicyDistributionClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.PolicyModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId});
            }

            localModelsWithSourceId.Where(policyModel => !policyModel.Items.Any()).ForEach(policyModel =>
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, policyModel.SourceId.Value);
                task2.Wait();

                policyModel.SourceId = null;
                policyModel.SourceTimestamp = null;

                var policyProfile = segment.PolicyProfiles.Single(profile => profile.Guid == policyModel.Guid);
                policyProfile.SourceId = null;
                policyProfile.SourceTimestamp = null;
                
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = policyModel.Name, SegmentId = segment.SourceId});
            });
        }

        
        private void InsertAndUpdatePolicyProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;

            var serverPolicyDistributionDictionary = GetServerPolicyProfiles(packageId, segmentId);

            segmentModel.PolicyModels.Where(policyModel => policyModel.Items.Any()).ForEach(policyModel =>
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorPolicyModel = policyModel.Map();

                if (!policyModel.SourceId.HasValue)
                {
                    InsertPolicyProfile(policyModel, packageId, segmentId, segment, collectorPolicyModel);
                }
                else
                {
                    var policyId = policyModel.SourceId.Value;

                    if (!serverPolicyDistributionDictionary.ContainsKey(policyId))
                    {
                        collectorPolicyModel.Id = 0;
                        InsertPolicyProfile(policyModel, packageId, segmentId, segment, collectorPolicyModel);
                        return;
                    }

                    if (collectorPolicyModel.IsEqualTo(serverPolicyDistributionDictionary[policyId])) return;
                    UpdatePolicyProfile(policyModel, packageId, segmentId, segment, collectorPolicyModel, policyId);
                }
            });
        }

        public Dictionary<long, PolicyDistributionModel> GetServerPolicyProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, PolicyDistributionModel>();

            var getAllTask = _collectorClient.PolicyDistributionClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverPolicyDistributions = getAllTask.Result;
            var serverPolicyDistributionDictionary =
                serverPolicyDistributions.ToDictionary(policyDistributionModel => policyDistributionModel.Id.Value);
            return serverPolicyDistributionDictionary;
        }

        private void UpdatePolicyProfile(PolicyModel policyModel, long packageId, long segmentId, ISegment segment,
            PolicyDistributionModel collectorPolicyModel, long policyId)
        {
            var task = _collectorClient.PolicyDistributionClient.PutAsync(packageId, segmentId, policyId, collectorPolicyModel);
            task.Wait();

            var policyProfile = segment.PolicyProfiles.Single(s => s.Guid == policyModel.Guid);
            policyModel.SourceTimestamp = policyProfile.SourceTimestamp = DateTime.Now;
            policyModel.IsDirty = policyProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update, Priority = 2, Name = policyModel.Name, SegmentId = segment.SourceId
            });
        }

        private void InsertPolicyProfile(PolicyModel policyModel, long packageId, long segmentId, ISegment segment,
            PolicyDistributionModel collectorPolicyModel)
        {
            var task = _collectorClient.PolicyDistributionClient.PostAsync(packageId, segmentId, collectorPolicyModel);
            task.Wait();

            var policyProfile = segment.PolicyProfiles.Single(s => s.Guid == policyModel.Guid);
            policyModel.SourceId = policyProfile.SourceId = task.Result;
            policyModel.SourceTimestamp = policyProfile.SourceTimestamp = DateTime.Now;
            policyModel.IsDirty = policyProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert, Priority = 2, Name = policyModel.Name, SegmentId = segment.SourceId
            });
        }

        private void DeleteStateProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.StateDistributionClient.GetAllAsync(packageId, segmentId);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.StateModels.Where(sm => sm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(sm => sm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var label = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = label, SegmentId = segment.SourceId});
            }

            localModelsWithSourceId.Where(stateModel => !stateModel.Items.Any()).ForEach(stateModel =>
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, stateModel.SourceId.Value);
                task2.Wait();

                stateModel.SourceId = null;
                stateModel.SourceTimestamp = null;

                var stateProfile = segment.StateProfiles.Single(profile => profile.Guid == stateModel.Guid);
                stateProfile.SourceId = null;
                stateProfile.SourceTimestamp = null;
                
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {
                    ActivityType = ActivityType.Delete, Priority = 2, Name = stateModel.Name, SegmentId = segment.SourceId
                });
            });
        }
        
        private void InsertAndUpdateStateProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverStateDistributionDictionary = GetServerStateProfiles(packageId, segmentId);

            segmentModel.StateModels.Where(stateModel => stateModel.Items.Any()).ForEach(stateModel =>
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorStateModel = stateModel.Map();

                if (!stateModel.SourceId.HasValue)
                {
                    InsertStateProfile(stateModel, packageId, segmentId, segment, collectorStateModel);
                }
                else
                {
                    var stateId = stateModel.SourceId.Value;

                    if (!serverStateDistributionDictionary.ContainsKey(stateId))
                    {
                        collectorStateModel.Id = 0;
                        InsertStateProfile(stateModel, packageId, segmentId, segment, collectorStateModel);
                        return;
                    }

                    if (collectorStateModel.IsEqualTo(serverStateDistributionDictionary[stateId])) return;
                    UpdateStateProfile(stateModel, packageId, segmentId, segment, collectorStateModel, stateId);
                }
            });
        }

        public Dictionary<long, StateDistributionModel> GetServerStateProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, StateDistributionModel>();

            var getAllTask = _collectorClient.StateDistributionClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverStateDistributions = getAllTask.Result;
            var serverStateDistributionDictionary =
                serverStateDistributions.ToDictionary(stateDistributionModel => stateDistributionModel.Id.Value);
            return serverStateDistributionDictionary;
        }

        private void UpdateStateProfile(StateModel stateModel, long packageId, long segmentId, ISegment segment,
            StateDistributionModel collectorStateModel, long stateId)
        {
            var task = _collectorClient.StateDistributionClient.PutAsync(packageId, segmentId, stateId, collectorStateModel);
            task.Wait();

            var stateProfile = segment.StateProfiles.Single(s => s.Guid == stateModel.Guid);
            stateModel.SourceTimestamp = stateProfile.SourceTimestamp = DateTime.Now;
            stateModel.IsDirty = stateProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update, Priority = 2, Name = stateModel.Name, SegmentId = segment.SourceId
            });
        }

        private void InsertStateProfile(StateModel stateModel, long packageId, long segmentId, ISegment segment,
            StateDistributionModel collectorStateModel)
        {
            var task = _collectorClient.StateDistributionClient.PostAsync(packageId, segmentId, collectorStateModel);
            task.Wait();

            var stateProfile = segment.StateProfiles.Single(s => s.Guid == stateModel.Guid);
            stateModel.SourceId = stateProfile.SourceId = task.Result;
            stateModel.SourceTimestamp = stateProfile.SourceTimestamp = DateTime.Now;
            stateModel.IsDirty = stateProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert, Priority = 2, Name = stateModel.Name, SegmentId = segment.SourceId
            });
        }


        private void DeleteProtectionClassProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.ProtectionClassDistributionClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.ProtectionClassModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId });
            }

            foreach (var protectionClassModel in localModelsWithSourceId.Where(protectionClassModel => !protectionClassModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, protectionClassModel.SourceId.Value);
                task2.Wait();

                protectionClassModel.SourceId = null;
                protectionClassModel.SourceTimestamp = null;

                var protectionClassProfile = segment.ProtectionClassProfiles.Single(profile => profile.Guid == protectionClassModel.Guid);
                protectionClassProfile.SourceId = null;
                protectionClassProfile.SourceTimestamp = null;

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = protectionClassModel.Name, SegmentId = segment.SourceId });
            }
        }
        
        private void InsertAndUpdateProtectionClassProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverProtectionClassDistributionDictionary = GetServerProtectionClassProfiles(packageId, segmentId);

            foreach (var protectionClassModel in segmentModel.ProtectionClassModels.Where(protectionClassModel => protectionClassModel.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorProtectionClassModel = protectionClassModel.Map();

                if (!protectionClassModel.SourceId.HasValue)
                {
                    InsertProtectionClassProfile(protectionClassModel, packageId, segmentId, segment, collectorProtectionClassModel);
                }
                else
                {
                    var sourceId = protectionClassModel.SourceId.Value;

                    if (!serverProtectionClassDistributionDictionary.ContainsKey(sourceId))
                    {
                        collectorProtectionClassModel.Id = 0;
                        InsertProtectionClassProfile(protectionClassModel, packageId, segmentId, segment, collectorProtectionClassModel);
                        continue;
                    }

                    if (collectorProtectionClassModel.IsEqualTo(serverProtectionClassDistributionDictionary[sourceId])) continue;
                    UpdateProtectionClassProfile(protectionClassModel, packageId, segmentId, segment, collectorProtectionClassModel, sourceId);
                }
            }
        }

        public Dictionary<long, ProtectionClassDistributionModel> GetServerProtectionClassProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, ProtectionClassDistributionModel>();

            var getAllTask = _collectorClient.ProtectionClassDistributionClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverProtectionClassDistributions = getAllTask.Result;
            var serverProtectionClassDistributionDictionary =
                serverProtectionClassDistributions.ToDictionary(stateDistributionModel => stateDistributionModel.Id.Value);
            return serverProtectionClassDistributionDictionary;
        }

        private void InsertProtectionClassProfile(ProtectionClassModel protectionClassModel, long packageId, long segmentId, ISegment segment,
            ProtectionClassDistributionModel collectorProtectionClassModel)
        {
            var task = _collectorClient.ProtectionClassDistributionClient.PostAsync(packageId, segmentId, collectorProtectionClassModel);
            task.Wait();

            var protectionClassProfile = segment.ProtectionClassProfiles.Single(s => s.Guid == protectionClassModel.Guid);
            protectionClassModel.SourceId = protectionClassProfile.SourceId = task.Result;
            protectionClassModel.SourceTimestamp = protectionClassProfile.SourceTimestamp = DateTime.Now;
            protectionClassModel.IsDirty = protectionClassProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert,
                Priority = 2,
                Name = protectionClassModel.Name,
                SegmentId = segment.SourceId
            });
        }

        private void UpdateProtectionClassProfile(ProtectionClassModel protectionClassModel, long packageId, long segmentId, ISegment segment,
            ProtectionClassDistributionModel collectorProtectionClassModel, long protectionClassId)
        {
            var task = _collectorClient.ProtectionClassDistributionClient.PutAsync(packageId, segmentId, protectionClassId, collectorProtectionClassModel);
            task.Wait();

            var protectionClassProfile = segment.ProtectionClassProfiles.Single(s => s.Guid == protectionClassModel.Guid);
            protectionClassModel.SourceTimestamp = protectionClassProfile.SourceTimestamp = DateTime.Now;
            protectionClassModel.IsDirty = protectionClassProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update,
                Priority = 2,
                Name = protectionClassModel.Name,
                SegmentId = segment.SourceId
            });
        }

        
        private void DeleteTotalInsuredValueProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.TotalInsuredValueDistributionClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.TotalInsuredValueModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId });
            }

            foreach (var totalInsuredValueModel in localModelsWithSourceId.Where(totalInsuredValueModel => !totalInsuredValueModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, totalInsuredValueModel.SourceId.Value);
                task2.Wait();

                totalInsuredValueModel.SourceId = null;
                totalInsuredValueModel.SourceTimestamp = null;

                var totalInsuredValueProfile = segment.TotalInsuredValueProfiles.Single(profile => profile.Guid == totalInsuredValueModel.Guid);
                totalInsuredValueProfile.SourceId = null;
                totalInsuredValueProfile.SourceTimestamp = null;

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = totalInsuredValueModel.Name, SegmentId = segment.SourceId });
            }
        }

        private void InsertAndUpdateTotalInsuredValueProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverTotalInsuredValueDistributionDictionary = GetServerTotalInsuredValueProfiles(packageId, segmentId);

            foreach (var totalInsuredValueModel in segmentModel.TotalInsuredValueModels.Where(totalInsuredValueModel => totalInsuredValueModel.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorTotalInsuredValueModel = totalInsuredValueModel.Map();

                if (!totalInsuredValueModel.SourceId.HasValue)
                {
                    InsertTotalInsuredValueProfile(totalInsuredValueModel, packageId, segmentId, segment, collectorTotalInsuredValueModel);
                }
                else
                {
                    var sourceId = totalInsuredValueModel.SourceId.Value;

                    if (!serverTotalInsuredValueDistributionDictionary.ContainsKey(sourceId))
                    {
                        collectorTotalInsuredValueModel.Id = 0;
                        InsertTotalInsuredValueProfile(totalInsuredValueModel, packageId, segmentId, segment, collectorTotalInsuredValueModel);
                        continue;
                    }

                    if (collectorTotalInsuredValueModel.IsEqualTo(serverTotalInsuredValueDistributionDictionary[sourceId])) continue;
                    UpdateTotalInsuredValueProfile(totalInsuredValueModel, packageId, segmentId, segment, collectorTotalInsuredValueModel, sourceId);
                }
            }
        }

        public Dictionary<long, TotalInsuredValueDistributionModel> GetServerTotalInsuredValueProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, TotalInsuredValueDistributionModel>();
            
            var getAllTask = _collectorClient.TotalInsuredValueDistributionClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverTotalInsuredValueDistributions = getAllTask.Result;
            var serverTotalInsuredValueDistributionDictionary =
                serverTotalInsuredValueDistributions.ToDictionary(stateDistributionModel => stateDistributionModel.Id.Value);
            return serverTotalInsuredValueDistributionDictionary;
        }

        private void InsertTotalInsuredValueProfile(TotalInsuredValueModel totalInsuredValueModel, long packageId, long segmentId, ISegment segment,
            TotalInsuredValueDistributionModel collectorTotalInsuredValueModel)
        {
            var task = _collectorClient.TotalInsuredValueDistributionClient.PostAsync(packageId, segmentId, collectorTotalInsuredValueModel);
            task.Wait();

            var totalInsuredValueProfile = segment.TotalInsuredValueProfiles.Single(s => s.Guid == totalInsuredValueModel.Guid);
            totalInsuredValueModel.SourceId = totalInsuredValueProfile.SourceId = task.Result;
            totalInsuredValueModel.SourceTimestamp = totalInsuredValueProfile.SourceTimestamp = DateTime.Now;
            totalInsuredValueModel.IsDirty = totalInsuredValueProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert,
                Priority = 2,
                Name = totalInsuredValueModel.Name,
                SegmentId = segment.SourceId
            });
        }

        private void UpdateTotalInsuredValueProfile(TotalInsuredValueModel totalInsuredValueModel, long packageId, long segmentId, ISegment segment,
            TotalInsuredValueDistributionModel collectorTotalInsuredValueModel, long totalInsuredValueId)
        {
            var task = _collectorClient.TotalInsuredValueDistributionClient.PutAsync(packageId, segmentId, totalInsuredValueId, collectorTotalInsuredValueModel);
            task.Wait();

            var totalInsuredValueProfile = segment.TotalInsuredValueProfiles.Single(s => s.Guid == totalInsuredValueModel.Guid);
            totalInsuredValueModel.SourceTimestamp = totalInsuredValueProfile.SourceTimestamp = DateTime.Now;
            totalInsuredValueModel.IsDirty = totalInsuredValueProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update,
                Priority = 2,
                Name = totalInsuredValueModel.Name,
                SegmentId = segment.SourceId
            });
        }



        private void DeleteConstructionTypeProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.ConstructionTypeDistributionClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.ConstructionTypeModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId });
            }

            foreach (var constructionTypeModel in localModelsWithSourceId.Where(constructionTypeModel => !constructionTypeModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, constructionTypeModel.SourceId.Value);
                task2.Wait();

                constructionTypeModel.SourceId = null;
                constructionTypeModel.SourceTimestamp = null;

                var constructionTypeProfile = segment.ConstructionTypeProfiles.Single(profile => profile.Guid == constructionTypeModel.Guid);
                constructionTypeProfile.SourceId = null;
                constructionTypeProfile.SourceTimestamp = null;

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = constructionTypeModel.Name, SegmentId = segment.SourceId });
            }
        }

        private void InsertAndUpdateConstructionTypeProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverConstructionTypeDistributionDictionary = GetServerConstructionTypeProfiles(packageId, segmentId);

            foreach (var constructionTypeModel in segmentModel.ConstructionTypeModels.Where(constructionTypeModel => constructionTypeModel.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorConstructionTypeModel = constructionTypeModel.Map();

                if (!constructionTypeModel.SourceId.HasValue)
                {
                    InsertConstructionTypeProfile(constructionTypeModel, packageId, segmentId, segment, collectorConstructionTypeModel);
                }
                else
                {
                    var sourceId = constructionTypeModel.SourceId.Value;

                    if (!serverConstructionTypeDistributionDictionary.ContainsKey(sourceId))
                    {
                        collectorConstructionTypeModel.Id = 0;
                        InsertConstructionTypeProfile(constructionTypeModel, packageId, segmentId, segment, collectorConstructionTypeModel);
                        continue;
                    }

                    if (collectorConstructionTypeModel.IsEqualTo(serverConstructionTypeDistributionDictionary[sourceId])) continue;
                    UpdateConstructionTypeProfile(constructionTypeModel, packageId, segmentId, segment, collectorConstructionTypeModel, sourceId);
                }
            }
        }

        public Dictionary<long, ConstructionTypeDistributionModel> GetServerConstructionTypeProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, ConstructionTypeDistributionModel>();

            var getAllTask = _collectorClient.ConstructionTypeDistributionClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverConstructionTypeDistributions = getAllTask.Result;
            var serverConstructionTypeDistributionDictionary =
                serverConstructionTypeDistributions.ToDictionary(stateDistributionModel => stateDistributionModel.Id.Value);
            return serverConstructionTypeDistributionDictionary;
        }

        private void InsertConstructionTypeProfile(ConstructionTypeModel constructionTypeModel, long packageId, long segmentId, ISegment segment,
            ConstructionTypeDistributionModel collectorConstructionTypeModel)
        {
            var task = _collectorClient.ConstructionTypeDistributionClient.PostAsync(packageId, segmentId, collectorConstructionTypeModel);
            task.Wait();

            var constructionTypeProfile = segment.ConstructionTypeProfiles.Single(s => s.Guid == constructionTypeModel.Guid);
            constructionTypeModel.SourceId = constructionTypeProfile.SourceId = task.Result;
            constructionTypeModel.SourceTimestamp = constructionTypeProfile.SourceTimestamp = DateTime.Now;
            constructionTypeModel.IsDirty = constructionTypeProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert,
                Priority = 2,
                Name = constructionTypeModel.Name,
                SegmentId = segment.SourceId
            });
        }

        private void UpdateConstructionTypeProfile(ConstructionTypeModel constructionTypeModel, long packageId, long segmentId, ISegment segment,
            ConstructionTypeDistributionModel collectorConstructionTypeModel, long constructionTypeId)
        {
            var task = _collectorClient.ConstructionTypeDistributionClient.PutAsync(packageId, segmentId, constructionTypeId, collectorConstructionTypeModel);
            task.Wait();

            var constructionTypeProfile = segment.ConstructionTypeProfiles.Single(s => s.Guid == constructionTypeModel.Guid);
            constructionTypeModel.SourceTimestamp = constructionTypeProfile.SourceTimestamp = DateTime.Now;
            constructionTypeModel.IsDirty = constructionTypeProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update,
                Priority = 2,
                Name = constructionTypeModel.Name,
                SegmentId = segment.SourceId
            });
        }


        private void DeleteOccupancyTypeProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.OccupancyTypeDistributionClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.OccupancyTypeModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId });
            }

            foreach (var occupancyTypeModel in localModelsWithSourceId.Where(occupancyTypeModel => !occupancyTypeModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, occupancyTypeModel.SourceId.Value);
                task2.Wait();

                occupancyTypeModel.SourceId = null;
                occupancyTypeModel.SourceTimestamp = null;

                var occupancyTypeProfile = segment.OccupancyTypeProfiles.Single(profile => profile.Guid == occupancyTypeModel.Guid);
                occupancyTypeProfile.SourceId = null;
                occupancyTypeProfile.SourceTimestamp = null;

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = occupancyTypeModel.Name, SegmentId = segment.SourceId });
            }
        }

        private void InsertAndUpdateOccupancyTypeProfiles(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverOccupancyTypeDistributionDictionary = GetServerOccupancyTypeProfiles(packageId, segmentId);

            foreach (var occupancyTypeModel in segmentModel.OccupancyTypeModels.Where(occupancyTypeModel => occupancyTypeModel.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorOccupancyTypeModel = occupancyTypeModel.Map();

                if (!occupancyTypeModel.SourceId.HasValue)
                {
                    InsertOccupancyTypeProfile(occupancyTypeModel, packageId, segmentId, segment, collectorOccupancyTypeModel);
                }
                else
                {
                    var sourceId = occupancyTypeModel.SourceId.Value;

                    if (!serverOccupancyTypeDistributionDictionary.ContainsKey(sourceId))
                    {
                        collectorOccupancyTypeModel.Id = 0;
                        InsertOccupancyTypeProfile(occupancyTypeModel, packageId, segmentId, segment, collectorOccupancyTypeModel);
                        continue;
                    }

                    if (collectorOccupancyTypeModel.IsEqualTo(serverOccupancyTypeDistributionDictionary[sourceId])) continue;
                    UpdateOccupancyTypeProfile(occupancyTypeModel, packageId, segmentId, segment, collectorOccupancyTypeModel, sourceId);
                }
            }
        }

        public Dictionary<long, OccupancyTypeDistributionModel> GetServerOccupancyTypeProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, OccupancyTypeDistributionModel>();

            var getAllTask = _collectorClient.OccupancyTypeDistributionClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverOccupancyTypeDistributions = getAllTask.Result;
            var serverOccupancyTypeDistributionDictionary =
                serverOccupancyTypeDistributions.ToDictionary(stateDistributionModel => stateDistributionModel.Id.Value);
            return serverOccupancyTypeDistributionDictionary;
        }

        private void InsertOccupancyTypeProfile(OccupancyTypeModel occupancyTypeModel, long packageId, long segmentId, ISegment segment,
            OccupancyTypeDistributionModel collectorOccupancyTypeModel)
        {
            var task = _collectorClient.OccupancyTypeDistributionClient.PostAsync(packageId, segmentId, collectorOccupancyTypeModel);
            task.Wait();

            var occupancyTypeProfile = segment.OccupancyTypeProfiles.Single(s => s.Guid == occupancyTypeModel.Guid);
            occupancyTypeModel.SourceId = occupancyTypeProfile.SourceId = task.Result;
            occupancyTypeModel.SourceTimestamp = occupancyTypeProfile.SourceTimestamp = DateTime.Now;
            occupancyTypeModel.IsDirty = occupancyTypeProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert,
                Priority = 2,
                Name = occupancyTypeModel.Name,
                SegmentId = segment.SourceId
            });
        }

        private void UpdateOccupancyTypeProfile(OccupancyTypeModel occupancyTypeModel, long packageId, long segmentId, ISegment segment,
            OccupancyTypeDistributionModel collectorOccupancyTypeModel, long occupancyTypeId)
        {
            var task = _collectorClient.OccupancyTypeDistributionClient.PutAsync(packageId, segmentId, occupancyTypeId, collectorOccupancyTypeModel);
            task.Wait();

            var occupancyTypeProfile = segment.OccupancyTypeProfiles.Single(s => s.Guid == occupancyTypeModel.Guid);
            occupancyTypeModel.SourceTimestamp = occupancyTypeProfile.SourceTimestamp = DateTime.Now;
            occupancyTypeModel.IsDirty = occupancyTypeProfile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update,
                Priority = 2,
                Name = occupancyTypeModel.Name,
                SegmentId = segment.SourceId
            });
        }



        
        private void DeleteWorkersCompStateAttachmentProfile(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.WorkersCompStateAttachmentDistributionsClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.WorkersCompStateAttachmentModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId });
            }

            foreach (var workersCompStateAttachmentModel in localModelsWithSourceId.Where(occupancyTypeModel => !occupancyTypeModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, workersCompStateAttachmentModel.SourceId.Value);
                task2.Wait();

                workersCompStateAttachmentModel.SourceId = null;
                workersCompStateAttachmentModel.SourceTimestamp = null;

                var workersCompStateAttachmentProfile = segment.WorkersCompStateAttachmentProfile;
                workersCompStateAttachmentProfile.SourceId = null;
                workersCompStateAttachmentProfile.SourceTimestamp = null;

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = workersCompStateAttachmentModel.Name, SegmentId = segment.SourceId });
            }
        }
        
        private void InsertAndUpdateWorkersCompStateAttachmentProfile(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverDistributionDictionary = GetServerWorkersCompStateAttachmentProfiles(packageId, segmentId);

            foreach (var workersCompStateAttachmentModel in segmentModel.WorkersCompStateAttachmentModels.Where(model => model.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var distributionModel = workersCompStateAttachmentModel.Map();

                if (!workersCompStateAttachmentModel.SourceId.HasValue)
                {
                    InsertWorkersCompStateAttachmentProfile(workersCompStateAttachmentModel, packageId, segmentId, segment, distributionModel);
                }
                else
                {
                    var sourceId = workersCompStateAttachmentModel.SourceId.Value;

                    if (!serverDistributionDictionary.ContainsKey(sourceId))
                    {
                        distributionModel.Id = 0;
                        InsertWorkersCompStateAttachmentProfile(workersCompStateAttachmentModel, packageId, segmentId, segment, distributionModel);
                        continue;
                    }

                    if (distributionModel.IsEqualTo(serverDistributionDictionary[sourceId])) continue;
                    UpdateWorkersCompStateAttachmentProfile(workersCompStateAttachmentModel, packageId, segmentId, segment, distributionModel, sourceId);
                }
            }
        }

        public Dictionary<long, WorkersCompStateAttachmentDistributionModel> GetServerWorkersCompStateAttachmentProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, WorkersCompStateAttachmentDistributionModel>();

            var getAllTask = _collectorClient.WorkersCompStateAttachmentDistributionsClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverDistributions = getAllTask.Result;
            var serverDistributionDictionary = serverDistributions.ToDictionary(distributionModel => distributionModel.Id.Value);
            return serverDistributionDictionary;
        }

        private void InsertWorkersCompStateAttachmentProfile(WorkersCompStateAttachmentModel model, 
            long packageId, long segmentId, ISegment segment, WorkersCompStateAttachmentDistributionModel collectorModel)
        {
            var task = _collectorClient.WorkersCompStateAttachmentDistributionsClient.PostAsync(packageId, segmentId, collectorModel);
            task.Wait();

            var profile = segment.WorkersCompStateAttachmentProfile;
            model.SourceId = profile.SourceId = task.Result;
            model.SourceTimestamp = profile.SourceTimestamp = DateTime.Now;
            model.IsDirty = profile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert,
                Priority = 2,
                Name = model.Name,
                SegmentId = segment.SourceId
            });
        }

        private void UpdateWorkersCompStateAttachmentProfile(WorkersCompStateAttachmentModel model, long packageId, long segmentId, ISegment segment,
            WorkersCompStateAttachmentDistributionModel collectorModel, long occupancyTypeId)
        {
            var task = _collectorClient.WorkersCompStateAttachmentDistributionsClient.PutAsync(packageId, segmentId, occupancyTypeId, collectorModel);
            task.Wait();

            var profile = segment.WorkersCompStateAttachmentProfile;
            model.SourceTimestamp = profile.SourceTimestamp = DateTime.Now;
            model.IsDirty = profile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update,
                Priority = 2,
                Name = model.Name,
                SegmentId = segment.SourceId
            });
        }


        
        private void DeleteWorkersCompStateHazardGroupProfile(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.WorkersCompStateHazardGroupDistributionsClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.WorkersCompStateHazardGroupModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId });
            }

            foreach (var model in localModelsWithSourceId.Where(localModel => !localModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, model.SourceId.Value);
                task2.Wait();

                model.SourceId = null;
                model.SourceTimestamp = null;

                var profile = segment.WorkersCompStateHazardGroupProfile;
                profile.SourceId = null;
                profile.SourceTimestamp = null;

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = model.Name, SegmentId = segment.SourceId });
            }
        }

        private void InsertAndUpdateWorkersCompStateHazardGroupProfile(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverDistributionDictionary = GetServerWorkersCompStateHazardGroupProfiles(packageId, segmentId);

            foreach (var stateHazardGroupModel in segmentModel.WorkersCompStateHazardGroupModels.Where(model => model.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var distributionModel = stateHazardGroupModel.Map();

                if (!stateHazardGroupModel.SourceId.HasValue)
                {
                    InsertWorkersCompStateHazardGroupProfile(stateHazardGroupModel, packageId, segmentId, segment, distributionModel);
                }
                else
                {
                    var sourceId = stateHazardGroupModel.SourceId.Value;

                    if (!serverDistributionDictionary.ContainsKey(sourceId))
                    {
                        distributionModel.Id = 0;
                        InsertWorkersCompStateHazardGroupProfile(stateHazardGroupModel, packageId, segmentId, segment, distributionModel);
                        continue;
                    }

                    if (distributionModel.IsEqualTo(serverDistributionDictionary[sourceId])) continue;
                    UpdateWorkersCompStateHazardGroupProfile(stateHazardGroupModel, packageId, segmentId, segment, distributionModel, sourceId);
                }
            }
        }

        public Dictionary<long, WorkersCompStateHazardGroupDistributionModel> GetServerWorkersCompStateHazardGroupProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, WorkersCompStateHazardGroupDistributionModel>();

            var getAllTask = _collectorClient.WorkersCompStateHazardGroupDistributionsClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverDistributions = getAllTask.Result;
            var serverDistributionDictionary = serverDistributions.ToDictionary(distributionModel => distributionModel.Id.Value);
            return serverDistributionDictionary;
        }

        private void InsertWorkersCompStateHazardGroupProfile(WorkersCompStateHazardGroupModel groupModel,
            long packageId, long segmentId, ISegment segment, WorkersCompStateHazardGroupDistributionModel collectorModel)
        {
            var task = _collectorClient.WorkersCompStateHazardGroupDistributionsClient.PostAsync(packageId, segmentId, collectorModel);
            task.Wait();

            var profile = segment.WorkersCompStateHazardGroupProfile;
            groupModel.SourceId = profile.SourceId = task.Result;
            groupModel.SourceTimestamp = profile.SourceTimestamp = DateTime.Now;
            groupModel.IsDirty = profile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert,
                Priority = 2,
                Name = groupModel.Name,
                SegmentId = segment.SourceId
            });
        }

        private void UpdateWorkersCompStateHazardGroupProfile(WorkersCompStateHazardGroupModel groupModel, long packageId, long segmentId, ISegment segment,
            WorkersCompStateHazardGroupDistributionModel collectorModel, long occupancyTypeId)
        {
            var task = _collectorClient.WorkersCompStateHazardGroupDistributionsClient.PutAsync(packageId, segmentId, occupancyTypeId, collectorModel);
            task.Wait();

            var profile = segment.WorkersCompStateHazardGroupProfile;
            groupModel.SourceTimestamp = profile.SourceTimestamp = DateTime.Now;
            groupModel.IsDirty = profile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update,
                Priority = 2,
                Name = groupModel.Name,
                SegmentId = segment.SourceId
            });
        }


        
        private void DeleteWorkersCompClassCodeProfile(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            
            var task = _collectorClient.WorkersCompStateClassCodeDistributionsClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.DistributionName);
            var localModelsWithSourceId = segmentModel.WorkersCompClassCodeModels.Where(pm => pm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(pm => pm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var name = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId });
            }

            foreach (var model in localModelsWithSourceId.Where(localModel => !localModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.DistributionClient.DeleteAsync(packageId, segmentId, model.SourceId.Value);
                task2.Wait();

                model.SourceId = null;
                model.SourceTimestamp = null;

                var profile = segment.WorkersCompClassCodeProfile;
                profile.SourceId = null;
                profile.SourceTimestamp = null;

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                { ActivityType = ActivityType.Delete, Priority = 2, Name = model.Name, SegmentId = segment.SourceId });
            }
        }

        private void InsertAndUpdateWorkersCompClassCodeProfile(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverDistributionDictionary = GetServerWorkersCompClassCodeProfiles(packageId, segmentId);

            foreach (var model in segmentModel.WorkersCompClassCodeModels.Where(model => model.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var distributionModel = model.Map();

                if (!model.SourceId.HasValue)
                {
                    InsertWorkersCompClassCodeProfile(model, packageId, segmentId, segment, distributionModel);
                }
                else
                {
                    var sourceId = model.SourceId.Value;

                    if (!serverDistributionDictionary.ContainsKey(sourceId))
                    {
                        distributionModel.Id = 0;
                        InsertWorkersCompClassCodeProfile(model, packageId, segmentId, segment, distributionModel);
                        continue;
                    }

                    if (distributionModel.IsEqualTo(serverDistributionDictionary[sourceId])) continue;
                    UpdateWorkersCompClassCodeProfile(model, packageId, segmentId, segment, distributionModel, sourceId);
                }
            }
        }

        
        public Dictionary<long, WorkersCompStateClassCodeDistributionModel> GetServerWorkersCompClassCodeProfiles(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, WorkersCompStateClassCodeDistributionModel>();

            var getAllTask = _collectorClient.WorkersCompStateClassCodeDistributionsClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverDistributions = getAllTask.Result;
            var serverDistributionDictionary = serverDistributions.ToDictionary(distributionModel => distributionModel.Id.Value);
            return serverDistributionDictionary;
        }

        
        private void InsertWorkersCompClassCodeProfile(WorkersCompClassCodeModel model,
            long packageId, long segmentId, ISegment segment, WorkersCompStateClassCodeDistributionModel collectorModel)
        {
            var task = _collectorClient.WorkersCompStateClassCodeDistributionsClient.PostAsync(packageId, segmentId, collectorModel);
            task.Wait();

            var profile = segment.WorkersCompClassCodeProfile;
            model.SourceId = profile.SourceId = task.Result;
            model.SourceTimestamp = profile.SourceTimestamp = DateTime.Now;
            model.IsDirty = profile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert,
                Priority = 2,
                Name = model.Name,
                SegmentId = segment.SourceId
            });
        }

        private void UpdateWorkersCompClassCodeProfile(WorkersCompClassCodeModel model, long packageId, long segmentId, ISegment segment,
            WorkersCompStateClassCodeDistributionModel collectorModel, long occupancyTypeId)
        {
            var task = _collectorClient.WorkersCompStateClassCodeDistributionsClient.PutAsync(packageId, segmentId, occupancyTypeId, collectorModel);
            task.Wait();

            var profile = segment.WorkersCompClassCodeProfile;
            model.SourceTimestamp = profile.SourceTimestamp = DateTime.Now;
            model.IsDirty = profile.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update,
                Priority = 2,
                Name = model.Name,
                SegmentId = segment.SourceId
            });
        }

        

        private void DeleteExposureSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.ExposureLossSetsClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.Name);
            var localModelsWithSourceId = segmentModel.ExposureSetModels.Where(esm => esm.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(esm => esm.SourceId.Value);
            foreach (var id in serverDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.ExposureLossSetsClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var label = serverDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = label, SegmentId = segment.SourceId});
            }

            localModelsWithSourceId.Where(exposureSetModel => !exposureSetModel.Items.Any()).ForEach(exposureSetModel =>
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.ExposureLossSetsClient.DeleteAsync(packageId, segmentId, exposureSetModel.SourceId.Value);
                task2.Wait();

                exposureSetModel.SourceId = null;
                exposureSetModel.SourceTimestamp = null;

                var exposureSet = segment.ExposureSets.Single(profile => profile.Guid == exposureSetModel.Guid);
                exposureSet.SourceId = null;
                exposureSet.SourceTimestamp = null;
                exposureSet.Ledger.Clear();
                
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {
                    ActivityType = ActivityType.Delete, Priority = 2, Name = exposureSetModel.Name, SegmentId = segment.SourceId
                });
            });
        }

        private void InsertAndUpdateExposureSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverExposureSetDictionary = GetServerExposureSets(packageId, segmentId);

            foreach (var exposureSetModel in segmentModel.ExposureSetModels.Where(exposureSetModel => exposureSetModel.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorExposureLossSetModel = exposureSetModel.Map();

                if (!exposureSetModel.SourceId.HasValue)
                {
                    InsertExposureSet(packageId, segmentId, exposureSetModel, segment, collectorExposureLossSetModel);
                }
                else
                {
                    var exposureSetId = exposureSetModel.SourceId.Value;

                    if (!serverExposureSetDictionary.ContainsKey(exposureSetId))
                    {
                        collectorExposureLossSetModel.Id = 0;
                        InsertExposureSet(packageId, segmentId, exposureSetModel, segment, collectorExposureLossSetModel);
                        continue;
                    }

                    if (collectorExposureLossSetModel.IsEqualTo(serverExposureSetDictionary[exposureSetId])) continue;
                    UpdateExposureSet(packageId, segmentId, exposureSetModel, segment, collectorExposureLossSetModel, exposureSetId);
                }
            }
        }

        public Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.ExposureSetModel> GetServerExposureSets(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.ExposureSetModel>();

            var getAllTask = _collectorClient.ExposureLossSetsClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverExposureSets = getAllTask.Result;
            var serverExposureSetDictionary = serverExposureSets.ToDictionary(exposureSetModel => exposureSetModel.Id.Value);
            return serverExposureSetDictionary;
        }

        private void UpdateExposureSet(long packageId, long segmentId, ExposureSetModel exposureSetModel, ISegment segment,
            ExposureSetModelPlus collectorExposureLossSetModel, long exposureSetId)
        {
            var task = _collectorClient.ExposureLossSetsClient.PutAsync(packageId, segmentId, exposureSetId, collectorExposureLossSetModel);
            task.Wait();

            var exposureSet = segment.ExposureSets.Single(s => s.Guid == exposureSetModel.Guid);
            exposureSetModel.SourceTimestamp = exposureSet.SourceTimestamp = DateTime.Now;
            exposureSetModel.IsDirty = exposureSet.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 2, Name = exposureSetModel.Name, SegmentId = segment.SourceId});
        }

        private void InsertExposureSet(long packageId, long segmentId, ExposureSetModel exposureSetModel, ISegment segment,
            ExposureSetModelPlus collectorExposureLossSetModel)
        {
            var task = _collectorClient.ExposureLossSetsClient.PostAsync(packageId, segmentId, collectorExposureLossSetModel);
            task.Wait();

            var exposureSet = segment.ExposureSets.Single(s => s.Guid == exposureSetModel.Guid);
            exposureSetModel.SourceId = exposureSet.SourceId = task.Result;
            exposureSetModel.SourceTimestamp = exposureSet.SourceTimestamp = DateTime.Now;
            exposureSetModel.IsDirty = exposureSet.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert, Priority = 2, Name = exposureSetModel.Name, SegmentId = segment.SourceId
            });
        }

        private void DeleteExposures(long packageId, SegmentModel segmentModel, ExposureSetModel exposureSetModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var exposureSetId = exposureSetModel.SourceId.Value;
            var name = $"{exposureSetModel.Name} Row";

            var task = _collectorClient.ExposureLossClient.GetAllAsync(packageId, segmentId, exposureSetId);
            task.Wait();
            var serverLossIds = task.Result.Select(x => x.Id.Value).ToList();
            var localLossIds = exposureSetModel.Items.Where(loss => loss.SourceId.HasValue).Select(loss => loss.SourceId.Value);
            foreach (var id in serverLossIds.Except(localLossIds))
            {
                var task2 = _collectorClient.ExposureLossClient.DeleteAsync(packageId, segmentId, exposureSetId, id);
                task2.Wait();

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 3, Name = name, SegmentId = segmentId});
            }
        }

        private void InsertAndUpdateExposures(long packageId, SegmentModel segmentModel, ExposureSetModel exposureSetModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var exposureSetId = exposureSetModel.SourceId.Value;

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            var exposureSet = segment.ExposureSets.Single(s => s.Guid == exposureSetModel.Guid);
            var serverExposuresDictionary = GetServerExposures(packageId, segmentId, exposureSetId);
            var name = $"{exposureSetModel.Name} Row";
            
            foreach (var exposure in exposureSetModel.Items)
            {
                var ledgerItem = exposureSet.Ledger[exposure.RowId];
                if (!exposure.SourceId.HasValue)
                {
                    InsertExposure(packageId, segmentId, exposureSetId, exposure, ledgerItem, name);
                }
                else
                {
                    var exposureId = exposure.SourceId.Value;

                    if (!serverExposuresDictionary.ContainsKey(exposureId))
                    {
                        exposure.Id = 0;
                        InsertExposure(packageId, segmentId, exposureSetId, exposure, ledgerItem, name);
                        continue;
                    }

                    if (exposure.IsEqualTo(serverExposuresDictionary[exposureId])) continue;
                    UpdateExposure(packageId, segmentId, exposureSetId, exposure, ledgerItem, exposureId, name);
                }
            }
        }

        public Dictionary<long, ExposureModel> GetServerExposures(long? packageId, long? segmentId, long? exposureSetId)
        {
            if (!packageId.HasValue || !segmentId.HasValue || !exposureSetId.HasValue)
                return new Dictionary<long, ExposureModel>();

            var getAllTask = _collectorClient.ExposureLossClient.GetAllAsync(packageId.Value, segmentId.Value, exposureSetId.Value);
            getAllTask.Wait();
            var serverExposures = getAllTask.Result;
            var serverExposuresDictionary = serverExposures.ToDictionary(exposure => exposure.Id.Value);
            return serverExposuresDictionary;
        }

        private void UpdateExposure(long packageId, long segmentId, long exposureSetId, ExposureModelPlus exposure, 
            ILedgerItem ledgerItem, long exposureId, string label)
        {
            var task = _collectorClient.ExposureLossClient.PutAsync(packageId, segmentId, exposureSetId, exposureId, exposure);
            task.Wait();

            exposure.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
            ledgerItem.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 3, Name = label, SegmentId = segmentId});
        }

        private void InsertExposure(long packageId, long segmentId, long exposureSetId, ExposureModelPlus exposure,
            ILedgerItem ledgerItem, string label)
        {
            var task = _collectorClient.ExposureLossClient.PostAsync(packageId, segmentId, exposureSetId, exposure);
            task.Wait();

            exposure.SourceId = ledgerItem.SourceId = task.Result;
            exposure.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
            ledgerItem.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Insert, Priority = 3, Name = label, SegmentId = segmentId});
        }

        private void DeleteAggregateLossSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.AggregateLossSetsClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverLossSetDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.Name);
            var localModelsWithSourceId =
                segmentModel.AggregateLossSetModels.Where(lossSetModel => lossSetModel.SourceId.HasValue).ToList();
            var localIds = localModelsWithSourceId.Select(lossSetModel => lossSetModel.SourceId.Value);
            foreach (var id in serverLossSetDictionary.Keys.Except(localIds))
            {
                var task2 = _collectorClient.AggregateLossSetsClient.DeleteAsync(packageId, segmentId, id);
                task2.Wait();

                var label = serverLossSetDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = label, SegmentId = segment.SourceId});
            }

            localModelsWithSourceId.Where(aggregateLossSetModel => !aggregateLossSetModel.Items.Any()).ForEach(aggregateLossSetModel =>
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.AggregateLossSetsClient.DeleteAsync(packageId, segmentId,
                    aggregateLossSetModel.SourceId.Value);
                task2.Wait();

                aggregateLossSetModel.SourceId = null;
                aggregateLossSetModel.SourceTimestamp = null;

                var lossSet = segment.AggregateLossSets.Single(aggregateLossSet => aggregateLossSet.Guid == aggregateLossSetModel.Guid);
                lossSet.SourceId = null;
                lossSet.SourceTimestamp = null;
                
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = aggregateLossSetModel.Name, SegmentId = segment.SourceId});
            });
        }

        private void InsertAndUpdateAggregateLossSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverAggregateLossSetDictionary = GetServerAggregateLossSets(packageId, segmentId);

            foreach (var lossSetModel in segmentModel.AggregateLossSetModels
                .Where(lossSetModel => lossSetModel.Items.Any()))
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorAggregateLossSetModel = lossSetModel.Map();
                if (!lossSetModel.SourceId.HasValue)
                {
                    InsertAggregateLossSet(lossSetModel, packageId, segmentId, segment, collectorAggregateLossSetModel);
                }
                else
                {
                    var lossSetId = lossSetModel.SourceId.Value;

                    if (!serverAggregateLossSetDictionary.ContainsKey(lossSetId))
                    {
                        collectorAggregateLossSetModel.Id = 0;
                        InsertAggregateLossSet(lossSetModel, packageId, segmentId, segment, collectorAggregateLossSetModel);
                        continue;
                    }

                    if (collectorAggregateLossSetModel.IsEqualTo(serverAggregateLossSetDictionary[lossSetId])) continue;
                    UpdateAggregateLossSet(lossSetModel, packageId, segmentId, segment, collectorAggregateLossSetModel, lossSetId);
                }
            }
        }

        public Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.AggregateLossSetModel> GetServerAggregateLossSets(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.AggregateLossSetModel>();

            var getAllTask = _collectorClient.AggregateLossSetsClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverAggregateLossSets = getAllTask.Result;
            var serverAggregateLossSetDictionary = serverAggregateLossSets.ToDictionary(lossSet => lossSet.Id.Value);
            return serverAggregateLossSetDictionary;
        }

        private void UpdateAggregateLossSet(AggregateLossSetModel lossSetModel, long packageId, long segmentId,
            ISegment segment, AggregateLossSetModelPlus collectorAggregateLossSetModel, long lossSetId)
        {
            var task = _collectorClient.AggregateLossSetsClient.PutAsync(packageId, segmentId, lossSetId, collectorAggregateLossSetModel);
            task.Wait();

            var lossSet = segment.AggregateLossSets.Single(s => s.Guid == lossSetModel.Guid);
            var name = lossSetModel.Name;

            lossSetModel.SourceTimestamp = lossSet.SourceTimestamp = DateTime.Now;
            lossSet.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 2, Name = name, SegmentId = segment.SourceId});
        }

        private void InsertAggregateLossSet(AggregateLossSetModel lossSetModel, long packageId, long segmentId,
            ISegment segment, AggregateLossSetModelPlus collectorAggregateLossSetModel)
        {
            var task = _collectorClient.AggregateLossSetsClient.PostAsync(packageId, segmentId, collectorAggregateLossSetModel);
            task.Wait();

            var lossSet = segment.AggregateLossSets.Single(s => s.Guid == lossSetModel.Guid);
            var name = lossSetModel.Name;

            lossSetModel.SourceId = lossSet.SourceId = task.Result;
            lossSetModel.SourceTimestamp = lossSet.SourceTimestamp = DateTime.Now;
            lossSetModel.IsDirty = lossSet.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Insert, Priority = 2, Name = name, SegmentId = segment.SourceId});
        }

        private void DeleteAggregateLosses(long packageId, SegmentModel segmentModel,
            AggregateLossSetModel aggregateLossSetModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var lossSetId = aggregateLossSetModel.SourceId.Value;
            var name = $"{aggregateLossSetModel.Name} Row";

            var task = _collectorClient.AggregateLossClient.GetAllAsync(packageId, segmentId, lossSetId);
            task.Wait();
            var serverLossIds = task.Result.Select(x => x.Id.Value).ToList();
            var localLossIds = aggregateLossSetModel.Items.Where(loss => loss.SourceId.HasValue).Select(loss => loss.SourceId.Value);
            foreach (var id in serverLossIds.Except(localLossIds))
            {
                var task2 = _collectorClient.AggregateLossClient.DeleteAsync(packageId, segmentId, lossSetId, id);
                task2.Wait();

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {
                    ActivityType = ActivityType.Delete, Priority = 3, Name = name, SegmentId = segmentId
                });
            }
        }

        private void InsertAndUpdateAggregateLosses(long packageId, SegmentModel segmentModel,
            AggregateLossSetModel aggregateLossSetModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var lossSetId = aggregateLossSetModel.SourceId.Value;

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            var aggregateLossSet = segment.AggregateLossSets.Single(s => s.Guid == aggregateLossSetModel.Guid);
            var name = $"{aggregateLossSetModel.Name} Row"; 
            
            var serverAggregateLossDictionary = GetServerAggregateLosses(packageId, segmentId, lossSetId);
            foreach (var loss in aggregateLossSetModel.Items)
            {
                var ledgerItem = aggregateLossSet.Ledger[loss.RowId];
                if (!loss.SourceId.HasValue)
                {
                    InsertAggregateLoss(packageId, segmentId, lossSetId, loss, ledgerItem, name);
                }
                else
                {
                    var lossId = loss.SourceId.Value;

                    if (!serverAggregateLossDictionary.ContainsKey(lossId))
                    {
                        loss.Id = 0;
                        InsertAggregateLoss(packageId, segmentId, lossSetId, loss, ledgerItem, name);
                        continue;
                    }

                    if (loss.IsEqualTo(serverAggregateLossDictionary[lossId])) continue;
                    UpdateAggregateLoss(packageId, segmentId, lossSetId, lossId, loss, ledgerItem, name);
                }
            }
        }

        public Dictionary<long, AggregateLossModel> GetServerAggregateLosses(long? packageId, long? segmentId, long? lossSetId)
        {
            if (!packageId.HasValue || !segmentId.HasValue || !lossSetId.HasValue)
                return new Dictionary<long, AggregateLossModel>();

            var getAllTask = _collectorClient.AggregateLossClient.GetAllAsync(packageId.Value, segmentId.Value, lossSetId.Value);
            getAllTask.Wait();
            var serverAggregateLosses = getAllTask.Result;
            var serverAggregateLossDictionary = serverAggregateLosses.ToDictionary(loss => loss.Id.Value);
            return serverAggregateLossDictionary;
        }

        private void UpdateAggregateLoss(long packageId, long segmentId, long lossSetId, long lossId, AggregateLossModelPlus loss,
            ILedgerItem ledgerItem, string label)
        {
            var task = _collectorClient.AggregateLossClient.PutAsync(packageId, segmentId, lossSetId, lossId, loss);
            task.Wait();

            loss.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
            ledgerItem.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 3, Name = label, SegmentId = segmentId});
        }

        private void InsertAggregateLoss(long packageId, long segmentId, long lossSetId, AggregateLossModelPlus loss,
            ILedgerItem ledgerItem, string label)
        {
            var task = _collectorClient.AggregateLossClient.PostAsync(packageId, segmentId, lossSetId, loss);
            task.Wait();

            var lossId = task.Result;
            loss.SourceId = ledgerItem.SourceId = lossId;
            loss.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
            ledgerItem.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Insert, Priority = 3, Name = label, SegmentId = segmentId});
        }

        private void DeleteIndividualLossSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.IndividualLossSetsClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverLossSetDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.Name);
            var localModelsWithSourceId =
                segmentModel.IndividualLossSetModels.Where(lossSetModel => lossSetModel.SourceId.HasValue).ToList();
            var localLossSetIds = localModelsWithSourceId.Select(lossSetModel => lossSetModel.SourceId.Value);
            foreach (var id in serverLossSetDictionary.Keys.Except(localLossSetIds))
            {
                var task2 = _collectorClient.IndividualLossSetsClient.DeleteAsync(packageId, segmentModel.SourceId.Value,
                    id);
                task2.Wait();

                var label = serverLossSetDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = label, SegmentId = segmentId});
            }

            foreach (var individualLossSetModel in localModelsWithSourceId.Where(individualLossSetModel =>
                !individualLossSetModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.IndividualLossSetsClient.DeleteAsync(packageId, segmentId,
                    individualLossSetModel.SourceId.Value);
                task2.Wait();

                individualLossSetModel.SourceId = null;
                individualLossSetModel.SourceTimestamp = null;
                var name = individualLossSetModel.Name;

                var lossSet = segment.IndividualLossSets.Single(individualLossSet => individualLossSet.Guid == individualLossSetModel.Guid);
                lossSet.SourceId = null;
                lossSet.SourceTimestamp = null;
                lossSet.Ledger.Clear();

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {
                    ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segment.SourceId
                });
            }
        }

        private void InsertAndUpdateIndividualLossSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverIndividualLossSetDictionary = GetServerIndividualLossSets(packageId, segmentId);

            segmentModel.IndividualLossSetModels.Where(lossSetModel => lossSetModel.Items.Any()).ForEach(lossSetModel =>
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorIndividualLossSetModel = lossSetModel.Map();
                if (!lossSetModel.SourceId.HasValue)
                {
                    InsertIndividualLossSet(lossSetModel, packageId, segmentId, segment, collectorIndividualLossSetModel);
                }
                else
                {
                    var lossSetId = lossSetModel.SourceId.Value;

                    if (!serverIndividualLossSetDictionary.ContainsKey(lossSetId))
                    {
                        collectorIndividualLossSetModel.Id = 0;
                        InsertIndividualLossSet(lossSetModel, packageId, segmentId, segment, collectorIndividualLossSetModel);
                        return;
                    }

                    if (collectorIndividualLossSetModel.IsEqualTo(serverIndividualLossSetDictionary[lossSetId])) return;
                    UpdateIndividualLossSet(lossSetModel, packageId, segmentId, segment, collectorIndividualLossSetModel, lossSetId);
                }
            });
        }

        public Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.IndividualLossSetModel> GetServerIndividualLossSets(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.IndividualLossSetModel>();

            var getAllTask = _collectorClient.IndividualLossSetsClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverIndividualLossSets = getAllTask.Result;
            var serverIndividualLossSetDictionary = serverIndividualLossSets.ToDictionary(lossSetModel => lossSetModel.Id.Value);
            return serverIndividualLossSetDictionary;
        }

        private void UpdateIndividualLossSet(IndividualLossSetModel lossSetModel, long packageId, long segmentId,
            ISegment segment, IndividualLossSetModelPlus collectorIndividualLossSetModel, long lossSetId)
        {
            var task = _collectorClient.IndividualLossSetsClient.PutAsync(packageId, segmentId, lossSetId, collectorIndividualLossSetModel);
            task.Wait();

            var lossSet = segment.IndividualLossSets.Single(s => s.Guid == lossSetModel.Guid);
            lossSetModel.SourceTimestamp = lossSet.SourceTimestamp = DateTime.Now;
            lossSet.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 2, Name = lossSet.Name, SegmentId = segment.SourceId});
        }

        private void InsertIndividualLossSet(IndividualLossSetModel lossSetModel, long packageId, long segmentId,
            ISegment segment, IndividualLossSetModelPlus collectorIndividualLossSetModel)
        {
            var task = _collectorClient.IndividualLossSetsClient.PostAsync(packageId, segmentId, collectorIndividualLossSetModel);
            task.Wait();

            var lossSet = segment.IndividualLossSets.Single(s => s.Guid == lossSetModel.Guid);
            lossSetModel.SourceId = lossSet.SourceId = task.Result;
            lossSetModel.SourceTimestamp = lossSet.SourceTimestamp = DateTime.Now;
            lossSet.IsDirty = false;

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Insert, Priority = 2, Name = lossSet.Name, SegmentId = segment.SourceId});
        }

        private void DeleteIndividualLosses(long packageId, SegmentModel segmentModel,
            IndividualLossSetModel individualLossSetModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var lossSetId = individualLossSetModel.SourceId.Value;

            var task = _collectorClient.IndividualLossClient.GetAllAsync(packageId, segmentId, lossSetId);
            task.Wait();

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            var individualLossSet = segment.IndividualLossSets.Single(s => s.Guid == individualLossSetModel.Guid);
            var name = $"{individualLossSetModel.Name} Row";
            var ledger = individualLossSet.Ledger;

            var serverIndividualLossIds = task.Result.Select(x => x.Id.Value).ToList();
            var localIndividualLossIds =
                individualLossSetModel.Items.Where(loss => loss.SourceId.HasValue).Select(alm => alm.SourceId.Value);
            var deleteIds = serverIndividualLossIds.Except(localIndividualLossIds);
            foreach (var id in deleteIds)
            {
                var task2 = _collectorClient.IndividualLossClient.DeleteAsync(packageId, segmentId, lossSetId, id);
                task2.Wait();

                ledger.FirstOrDefault(item => item.SourceId == id)?.Clear();
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {
                    ActivityType = ActivityType.Delete, Priority = 3, Name = name, SegmentId = segmentId
                });
            }
        }

        private void InsertAndUpdateIndividualLosses(long packageId, SegmentModel segmentModel,
            IndividualLossSetModel individualLossSetModel, BackgroundWorker worker)
        {
            var segmentId = segmentModel.SourceId.Value;
            var lossSetId = individualLossSetModel.SourceId.Value;

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            var individualLossSet = segment.IndividualLossSets.Single(s => s.Guid == individualLossSetModel.Guid);
            var serverIndividualLossDictionary = GetServerIndividualLosses(packageId, segmentId, lossSetId);
            if (!individualLossSetModel.Items.Any()) return;

            var name = $"{individualLossSetModel.Name} Row";

            var lossCount = individualLossSetModel.Items.Count;
            const int pageSize = 100;
            var pages = Convert.ToInt32(Math.Ceiling(lossCount / (double) pageSize));
            for (var pageNumber = 0; pageNumber < pages; pageNumber++)
            {
                var itemsSubset = individualLossSetModel.Items.Skip(pageNumber * pageSize).Take(pageSize);
                var first = pageNumber * pageSize + 1;
                var last = Math.Min(first + pageSize - 1, lossCount);
                worker?.ReportProgress(0, $"Uploading {BexConstants.SegmentName.ToLower()} <{segmentModel.Name}>" +
                                          "\r\n" +
                                          $"{individualLossSetModel.Name} " +
                                          $"({first:N0}-{last:N0} of {individualLossSetModel.Items.Count:N0}) ...");

                var tasks = new List<Task>();
                foreach (var loss in itemsSubset)
                {
                    var ledgerItem = individualLossSet.Ledger[loss.RowId];
                    if (!loss.SourceId.HasValue)
                    {
                        var task = InsertIndividualLoss(packageId, segmentId, lossSetId, loss, ledgerItem, name);
                        tasks.Add(task);
                    }
                    else
                    {
                        var lossId = loss.SourceId.Value;
                        Task task;
                        if (!serverIndividualLossDictionary.ContainsKey(lossId))
                        {
                            loss.Id = 0;
                            task = InsertIndividualLoss(packageId, segmentId, lossSetId, loss, ledgerItem, name);
                            tasks.Add(task);
                        }
                        else
                        {
                            if (loss.IsEqualTo(serverIndividualLossDictionary[lossId])) continue;
                            task = UpdateIndividualLoss(packageId, segmentId, lossSetId, loss, ledgerItem, lossId, name);
                        }

                        tasks.Add(task);
                    }
                }

                Task.WaitAll(tasks.ToArray());
            }
        }

        public Dictionary<long, IndividualLossModel> GetServerIndividualLosses(long? packageId, long? segmentId,
            long? lossSetId)
        {
            if (!packageId.HasValue || !segmentId.HasValue || !lossSetId.HasValue)
                return new Dictionary<long, IndividualLossModel>();

            var getAllTask = _collectorClient.IndividualLossClient.GetAllAsync(packageId.Value, segmentId.Value, lossSetId.Value);
            getAllTask.Wait();
            var serverIndividualLosses = getAllTask.Result;
            var serverIndividualLossDictionary = serverIndividualLosses.ToDictionary(loss => loss.Id.Value);
            return serverIndividualLossDictionary;
        }

        private Task UpdateIndividualLoss(long packageId, long segmentId, long lossSetId, IndividualLossModelPlus loss,
            ILedgerItem ledgerItem, long lossId, string label)
        {
            var task = _collectorClient.IndividualLossClient.PutAsync(packageId, segmentId, lossSetId, lossId, loss)
                .ContinueWith(task1 =>
                {
                    if (!task1.IsFaulted)
                    {
                        if (!loss.SourceId.HasValue) loss.SourceId = ledgerItem.SourceId = lossId;
                        loss.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
                        ledgerItem.IsDirty = false;
                    }
                    else
                    {
                        throw new ApplicationException($"{BexConstants.IndividualLossSetName.ToStartOfSentence()} API Failed",
                            task1.Exception);
                    }
                });


            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Update, Priority = 3, Name = label, SegmentId = segmentId
            });
            return task;
        }

        private Task InsertIndividualLoss(long packageId, long segmentId, long lossSetId, IndividualLossModelPlus loss,
            ILedgerItem ledgerItem, string label)
        {
            var task = _collectorClient.IndividualLossClient.PostAsync(packageId, segmentId, lossSetId, loss).ContinueWith(task1 =>
            {
                if (!task1.IsFaulted)
                {
                    var lossId = task1.Result;
                    loss.SourceId = ledgerItem.SourceId = lossId;
                    loss.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
                    ledgerItem.IsDirty = false;
                }
                else
                {
                    throw new ApplicationException($"{BexConstants.IndividualLossSetName.ToStartOfSentence()} API Failed",
                        task1.Exception);
                }
            });

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert, Priority = 3, Name = label, SegmentId = segmentId
            });

            return task;
        }

        private void InsertAndUpdateRateChangeSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var serverRateChangeSetDictionary = GetServerRateChangeSets(packageId, segmentId);

            segmentModel.RateChangeSetModels.Where(rateChangeSetModel => rateChangeSetModel.Items.Any()).ForEach(rateChangeSetModel =>
            {
                var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
                var collectorRateChangeSetModel = rateChangeSetModel.Map();
                if (!rateChangeSetModel.SourceId.HasValue)
                {
                    InsertRateChangeSet(rateChangeSetModel, packageId, segmentId, segment, collectorRateChangeSetModel);
                }
                else
                {
                    var rateChangeSetId = rateChangeSetModel.SourceId.Value;

                    if (!serverRateChangeSetDictionary.ContainsKey(rateChangeSetId))
                    {
                        collectorRateChangeSetModel.Id = 0;
                        InsertRateChangeSet(rateChangeSetModel, packageId, segmentId, segment, collectorRateChangeSetModel);
                        return;
                    }

                    if (collectorRateChangeSetModel.IsEqualTo(serverRateChangeSetDictionary[rateChangeSetId])) return;
                    UpdateRateChangeSet(rateChangeSetModel, packageId, segmentId, segment, collectorRateChangeSetModel, rateChangeSetId);
                }
            });
        }

        public Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.RateChangeSetModel> GetServerRateChangeSets(long? packageId, long? segmentId)
        {
            if (!packageId.HasValue || !segmentId.HasValue) return new Dictionary<long, MunichRe.Bex.ApiClient.CollectorApi.RateChangeSetModel>();

            var getAllTask = _collectorClient.RateChangeSetsClient.GetAllAsync(packageId.Value, segmentId.Value);
            getAllTask.Wait();
            var serverRateChangeSets = getAllTask.Result;
            var serverRateChangeSetDictionary = serverRateChangeSets.ToDictionary(lossSetModel => lossSetModel.Id.Value);
            return serverRateChangeSetDictionary;
        }

        private void DeleteRateChangeSets(long packageId, SegmentModel segmentModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);

            var task = _collectorClient.RateChangeSetsClient.GetAllAsync(packageId, segmentModel.SourceId.Value);
            task.Wait();

            var serverRateChangeSetDictionary = task.Result.ToDictionary(k => k.Id.Value, v => v.Name);
            var localModelsWithSourceId = segmentModel.RateChangeSetModels.Where(rateChangeSetModel => rateChangeSetModel.SourceId.HasValue)
                .ToList();
            var localRateChangeSetIds = localModelsWithSourceId.Select(rateChangeSetModel => rateChangeSetModel.SourceId.Value);
            foreach (var id in serverRateChangeSetDictionary.Keys.Except(localRateChangeSetIds))
            {
                var task2 = _collectorClient.RateChangeSetsClient.DeleteAsync(packageId, segmentModel.SourceId.Value, id);
                task2.Wait();

                var name = serverRateChangeSetDictionary[id];
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = name, SegmentId = segmentId});
            }

            foreach (var rateChangeSetModel in localModelsWithSourceId.Where(rateChangeSetModel => !rateChangeSetModel.Items.Any()))
            {
                //unlike other deletes, this exists in workbook but need to deleted from server because it's row-less
                var task2 = _collectorClient.RateChangeSetsClient.DeleteAsync(packageId, segmentId,
                    rateChangeSetModel.SourceId.Value);
                task2.Wait();

                rateChangeSetModel.SourceId = null;
                rateChangeSetModel.SourceTimestamp = null;

                var rateChangeSet = segment.RateChangeSets.Single(rcSet => rcSet.Guid == rateChangeSetModel.Guid);
                rateChangeSet.SourceId = null;
                rateChangeSet.SourceTimestamp = null;
                rateChangeSet.Ledger.Clear();

                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 2, Name = rateChangeSetModel.Name, SegmentId = segment.SourceId});
            }
        }

        private void InsertRateChangeSet(RateChangeSetModel rateChangeSetModel, long packageId, long segmentId,
            ISegment segment, MunichRe.Bex.ApiClient.CollectorApi.RateChangeSetModel collectorRateChangeSetModel)
        {
            var task = _collectorClient.RateChangeSetsClient.PostAsync(packageId, segmentId, collectorRateChangeSetModel);
            task.Wait();

            var rateChangeSet = segment.RateChangeSets.Single(s => s.Guid == rateChangeSetModel.Guid);
            rateChangeSetModel.SourceId = rateChangeSet.SourceId = task.Result;
            rateChangeSetModel.SourceTimestamp = rateChangeSet.SourceTimestamp = DateTime.Now;
            rateChangeSet.IsDirty = false;

            var name = rateChangeSetModel.Name;
            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Insert, Priority = 2, Name = name, SegmentId = segment.SourceId});
        }

        private void UpdateRateChangeSet(RateChangeSetModel rateChangeSetModel, long packageId, long segmentId,
            ISegment segment, MunichRe.Bex.ApiClient.CollectorApi.RateChangeSetModel collectorRateChangeSetModel, long lossSetId)
        {
            var task = _collectorClient.RateChangeSetsClient.PutAsync(packageId, segmentId, lossSetId, collectorRateChangeSetModel);
            task.Wait();

            var rateChangeSet = segment.RateChangeSets.Single(s => s.Guid == rateChangeSetModel.Guid);
            rateChangeSetModel.SourceTimestamp = rateChangeSet.SourceTimestamp = DateTime.Now;
            rateChangeSet.IsDirty = false;

            var name = rateChangeSetModel.Name;
            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 2, Name = name, SegmentId = segment.SourceId});
        }


        private void DeleteRateChanges(long packageId, SegmentModel segmentModel,
            RateChangeSetModel rateChangeSetModel)
        {
            var segmentId = segmentModel.SourceId.Value;
            var rateChangeSetId = rateChangeSetModel.SourceId.Value;

            var task = _collectorClient.RateChangeClient.GetAllAsync(packageId, segmentId, rateChangeSetId);
            task.Wait();

            var serverRateChangeIds = task.Result.Select(x => x.Id.Value).ToList();
            var localRateChangeIds = rateChangeSetModel.Items.Where(loss => loss.SourceId.HasValue).Select(alm => alm.SourceId.Value);
            var deleteIds = serverRateChangeIds.Except(localRateChangeIds);

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            var rateChangeSet = segment.RateChangeSets.Single(s => s.Guid == rateChangeSetModel.Guid);
            var ledger = rateChangeSet.Ledger;
            var name = $"{rateChangeSetModel.Name} Row";

            foreach (var id in deleteIds)
            {
                var task2 = _collectorClient.RateChangeClient.DeleteAsync(packageId, segmentId, rateChangeSetId, id);
                task2.Wait();

                ledger.FirstOrDefault(item => item.SourceId == id)?.Clear();
                _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                    {ActivityType = ActivityType.Delete, Priority = 3, Name = name, SegmentId = segmentId});
            }
        }

        private void InsertAndUpdateRateChanges(long packageId, SegmentModel segmentModel,
            RateChangeSetModel rateChangeSetModel, BackgroundWorker worker)
        {
            var segmentId = segmentModel.SourceId.Value;
            var rateChangeSetId = rateChangeSetModel.SourceId.Value;

            var segment = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(s => s.Guid == segmentModel.Guid);
            var rateChangeSet = segment.RateChangeSets.Single(s => s.Guid == rateChangeSetModel.Guid);
            var serverRateChangeDictionary = GetServerRateChanges(packageId, segmentId, rateChangeSetId);
            if (!rateChangeSetModel.Items.Any()) return;

            var rateChangeCount = rateChangeSetModel.Items.Count;
            const int pageSize = 100;
            var pages = Convert.ToInt32(Math.Ceiling(rateChangeCount / (double) pageSize));
            for (var pageNumber = 0; pageNumber < pages; pageNumber++)
            {
                var itemsSubset = rateChangeSetModel.Items.Skip(pageNumber * pageSize).Take(pageSize);
                var first = pageNumber * pageSize + 1;
                var last = Math.Min(first + pageSize - 1, rateChangeCount);
                worker?.ReportProgress(0, $"Uploading {BexConstants.SegmentName.ToLower()} <{segmentModel.Name}>" +
                                          "\r\n" +
                                          $"{rateChangeSetModel.Name} " +
                                          $"({first:N0}-{last:N0} of {rateChangeSetModel.Items.Count:N0}) ...");

                var name = $"{rateChangeSetModel.Name} Row";
                var tasks = new List<Task>();
                foreach (var rateChange in itemsSubset)
                {
                    var ledgerItem = rateChangeSet.Ledger[rateChange.RowId];
                    if (!rateChange.SourceId.HasValue)
                    {
                        var task = InsertRateChange(packageId, segmentId, rateChangeSetId, rateChange, ledgerItem, name);
                        tasks.Add(task);
                    }
                    else
                    {
                        var rateChangeId = rateChange.SourceId.Value;
                        Task task;
                        if (!serverRateChangeDictionary.ContainsKey(rateChangeId))
                        {
                            rateChange.Id = 0;
                            task = InsertRateChange(packageId, segmentId, rateChangeSetId, rateChange, ledgerItem, name);
                            tasks.Add(task);
                        }
                        else
                        {
                            if (rateChange.IsEqualTo(serverRateChangeDictionary[rateChangeId])) continue;
                            task = UpdateRateChange(packageId, segmentId, rateChangeSetId, rateChange, ledgerItem, rateChangeId, name);
                        }

                        tasks.Add(task);
                    }
                }

                Task.WaitAll(tasks.ToArray());
            }
        }


        private Task InsertRateChange(long packageId, long segmentId, long rateChangeSetId, RateChangeModelPlus rateChange,
            ILedgerItem ledgerItem, string label)
        {
            var task = _collectorClient.RateChangeClient.PostAsync(packageId, segmentId, rateChangeSetId, rateChange).ContinueWith(task1 =>
            {
                if (!task1.IsFaulted)
                {
                    var lossId = task1.Result;
                    rateChange.SourceId = ledgerItem.SourceId = lossId;
                    rateChange.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
                    ledgerItem.IsDirty = false;
                }
                else
                {
                    throw new ApplicationException($"{BexConstants.RateChangeSetName.ToStartOfSentence()} API Failed",
                        task1.Exception);
                }
            });

            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
            {
                ActivityType = ActivityType.Insert, Priority = 3, Name = label, SegmentId = segmentId
            });

            return task;
        }

        private Task UpdateRateChange(long packageId, long segmentId, long rateChangeSetId, RateChangeModelPlus rateChange,
            ILedgerItem ledgerItem, long lossId, string label)
        {
            var task = _collectorClient.RateChangeClient.PutAsync(packageId, segmentId, rateChangeSetId, lossId, rateChange)
                .ContinueWith(task1 =>
                {
                    if (!task1.IsFaulted)
                    {
                        if (!rateChange.SourceId.HasValue) rateChange.SourceId = ledgerItem.SourceId = lossId;
                        rateChange.SourceTimestamp = ledgerItem.SourceTimestamp = DateTime.Now;
                        ledgerItem.IsDirty = false;
                    }
                    else
                    {
                        throw new ApplicationException($"{BexConstants.RateChangeSetName.ToStartOfSentence()} API Failed",
                            task1.Exception);
                    }
                });


            _activityTracker.UnitsOfWork.AddNew(new ActivityUnitOfWork
                {ActivityType = ActivityType.Update, Priority = 3, Name = label, SegmentId = segmentId});
            return task;
        }

        public Dictionary<long, RateChangeModel> GetServerRateChanges(long? packageId, long? segmentId, long? rateChangeSetId)
        {
            if (!packageId.HasValue || !segmentId.HasValue || !rateChangeSetId.HasValue)
                return new Dictionary<long, RateChangeModel>();

            var getAllTask = _collectorClient.RateChangeClient.GetAllAsync(packageId.Value, segmentId.Value, rateChangeSetId.Value);
            getAllTask.Wait();
            var serverRateChanges = getAllTask.Result;
            var serverRateChangeDictionary = serverRateChanges.ToDictionary(loss => loss.Id.Value);
            return serverRateChangeDictionary;
        }
    }
}