using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Exceptions;
using PionlearClient.KeyDataFolder;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelUtilities.RangeSizeModifier;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.Properties;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;
using IndividualLossModel = MunichRe.Bex.ApiClient.CollectorApi.IndividualLossModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class WorkbookRebuilder
    {
        private readonly BexCommunicationManager _bexCommunicationManager;
        private readonly IWorkbookLogger _logger;
        private readonly IDictionary<int, ISubline> _allSublines;
        
        public WorkbookRebuilder(BexCommunicationManager bexCommunicationManager, IWorkbookLogger logger)
        {
            _bexCommunicationManager = bexCommunicationManager;
            _logger = logger;

            _allSublines = typeof(BaseSubline)
                .Assembly.GetTypes()
                .Where(t => t.IsSubclassOf(typeof(BaseSubline)) && !t.IsAbstract)
                .Select(t => (ISubline) Activator.CreateInstance(t))
                .ToDictionary(subline => subline.Code);
        }
        
        internal void RebuildWithProgress(IPackage package, long serverPackageId)
        {
            try
            {
                #region parameterize form
                var marqueeProgressBarViewModel = new MarqueeProgressBarViewModel();
                var marqueeProgressBar = new MarqueeProgressBar(marqueeProgressBarViewModel);
                var form = new MarqueeProgressForm(marqueeProgressBar)
                {
                    Text = BexConstants.ApplicationName,
                    Height = (int)FormSizeHeight.Medium,
                    Width = (int)FormSizeWidth.Medium,
                    StartPosition = FormStartPosition.CenterScreen,
                    ControlBox = false,
                };
                #endregion

                var backgroundWorker = new BackgroundWorker { WorkerReportsProgress = true };
                var arguments = new List<object> {package, serverPackageId};
                backgroundWorker.DoWork += RebuildWithFrozenExcel;

                backgroundWorker.ProgressChanged += (sender, e) =>
                {
                    var status = e.UserState.ToString();
                    marqueeProgressBarViewModel.Status = status;
                };

                backgroundWorker.RunWorkerCompleted += (sender, e) =>
                {
                    if (e.Error == null)
                    {
                        marqueeProgressBarViewModel.Message = $"{BexConstants.RebuildName} workbook successful";
                        marqueeProgressBarViewModel.Image = Resources.Success.ToBitmapSource();
                    }
                    else
                    {
                        var startMessage = $"{BexConstants.RebuildName} workbook failed";
                        if (e.Error == null)
                        {
                            marqueeProgressBarViewModel.Message = startMessage;
                        }
                        else
                        {
                            var message = string.Join(" ", e.Error.GetInnerExceptions().Select(x => x.Message).ToArray());
                            marqueeProgressBarViewModel.Message = $"{startMessage}{Environment.NewLine}{message}";
                        }

                        marqueeProgressBarViewModel.Image = Resources.Stop.ToBitmapSource();
                    }

                    form.ControlBox = true;
                    var ranSuccessfully = e.Error == null;
                    if (ranSuccessfully)
                    {
                        marqueeProgressBarViewModel.SetSuccessAppearance();
                        form.Height = (int)FormSizeHeight.Small;
                    }
                    else
                    {
                        marqueeProgressBarViewModel.SetFailureAppearance();

                        var bp = MessageHelper.GetBoxProperties(marqueeProgressBarViewModel.Message);
                        form.Height = (int)bp.Height;
                        form.Width = (int)bp.Width;
                    }
                };

                backgroundWorker.RunWorkerAsync(arguments);
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                _logger.WriteNew(ex);
                const string message = "Rebuild workbook failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private void RebuildWithFrozenExcel(object sender, DoWorkEventArgs e)
        {
            using (new ExcelScreenUpdateDisabler())
            {
                Rebuild(sender, e);
            }
        }

        private void Rebuild(object sender, DoWorkEventArgs e)
        {
            var genericList = e.Argument as List<object>;
            Debug.Assert(genericList != null, "Rebuild arguments not supplied");
            var package = (IPackage) genericList[0];
            var serverPackageId = (long)genericList[1];

            const int stringCounterItemMaximum = 5;
            var worker = sender as BackgroundWorker;
            
            worker?.ReportProgress(0, $"Starting {BexConstants.RebuildName.ToLower()} ...");
            var serverPackage = _bexCommunicationManager.TryToGetSubmissionPackageModel(serverPackageId);
            if (serverPackage == null)
            {
                throw new NoPackageException(serverPackageId);
            }

            var sb = new StringBuilder();
            sb.AppendLine($"{BexConstants.RebuildName}ing {BexConstants.PackageName.ToLower()} <{serverPackage.Name}> ...");
            worker?.ReportProgress(0, sb);
            MapServerPackageToWorkbook(serverPackage, package);
            var serverSegments = _bexCommunicationManager.GetServerSegments(serverPackageId);

            foreach (var serverSegmentPair in serverSegments)
            {
                MarkItemsAsDone(sb);
                sb.AppendLine($"{BexConstants.RebuildName}ing { BexConstants.SegmentName.ToLower()} <{ serverSegmentPair.Value.Name}> ...");
                if (CountStringItems(sb) > stringCounterItemMaximum) RemoveFirstRow(sb);
                worker?.ReportProgress(0, sb);
                
                var segment = new Segment(isNotCalledFromJson: true) {SourceId = serverSegmentPair.Key};
                var serverSegment = serverSegmentPair.Value;
                segment.IsWorkerCompClassCodeActive = _bexCommunicationManager.GetServerWorkersCompClassCodeProfiles(serverPackageId, serverSegment.Id).Any();

                var sublineAllocations = serverSegment.ValidSublines;
                var sublines = sublineAllocations.Select(sa => _allSublines[Convert.ToInt32(sa.Id)]).ToList();

                var sublinesCopy = new List<ISubline>();
                foreach (var subline in sublines)
                {
                    var sublineCopy = subline.DeepClone();
                    sublineCopy.SegmentId = segment.Id;
                    sublinesCopy.Add(sublineCopy);
                }
                segment.UpdateSublines(sublinesCopy);

                var serverPolicyProfiles = CreatePolicyProfiles(serverPackageId, serverSegment, segment);
                var serverStateProfiles = CreateStateProfiles(serverPackageId, serverSegment, segment);
                var serverHazardProfiles = CreateHazardProfiles(serverPackageId, serverSegment, segment);
            
                var serverConstructionTypeProfiles = CreateConstructionTypeProfiles(serverPackageId, serverSegment, segment);
                var serverOccupancyTypeProfiles = CreateOccupancyTypeProfiles(serverPackageId, serverSegment, segment);
                var serverProtectionClassProfiles = CreateProtectionClassProfiles(serverPackageId, serverSegment, segment);
                var serverTotalInsuredValueProfiles = CreateTotalInsuredValueProfiles(serverPackageId, serverSegment, segment);

                var serverWorkersCompStateHazardGroupProfiles = CreateWorkersCompStateHazardGroupProfiles(serverPackageId, serverSegment, segment);
                var serverWorkersCompClassCodeProfiles = CreateWorkersCompClassCodeProfiles(serverPackageId, serverSegment, segment).ToList();
                var serverWorkersCompRetentionProfiles = CreateWorkersCompRetentionProfiles(serverPackageId, serverSegment, segment);
                
                var serverExposureSets = CreateExposureSets(serverPackageId, serverSegment, segment);
                var serverAggregateLossSets = CreateAggregateLossSets(serverPackageId, serverSegment, segment);
                var serverIndividualLossSets = CreateIndividualLossSets(serverPackageId, serverSegment, segment);
                var serverRateChangeSets = CreateRateChangeSets(serverPackageId, serverSegment, segment);

                package.Segments.Add(segment);
                using (new ExcelEventDisabler())
                {
                    segment.WorksheetManager.CreateWorksheet();
                    segment.IsSelected = true;
                }
                segment.WorksheetManager.Worksheet = segment.Name.GetWorksheet();

                MapServerSegmentToSegment(serverPackageId, serverSegment, segment);
                
                PopulateProspectiveExposure(segment, serverSegmentPair);
                PopulateSublines(serverSegmentPair, segment);
                PopulateUmbrellaTypes(serverSegmentPair, segment);

                PopulatePolicyProfiles(segment, serverPolicyProfiles);
                PopulateStateProfiles(segment, serverStateProfiles);
                PopulateHazardProfiles(segment, serverHazardProfiles);

                PopulateConstructionTypeProfiles(segment, serverConstructionTypeProfiles);
                PopulateOccupancyTypeProfiles(segment, serverOccupancyTypeProfiles);
                PopulateProtectionClassProfiles(segment, serverProtectionClassProfiles);
                PopulateTotalInsuredValueProfiles(segment, serverTotalInsuredValueProfiles);

                PopulateMinnesotaRetention(segment.WorkersCompMinnesotaRetention, serverSegment.MinnesotaRetentionId);
                PopulateWorkersCompStateHazardGroupProfile(segment, serverWorkersCompStateHazardGroupProfiles);
                PopulateWorkersCompClassCodeProfile(segment, serverWorkersCompClassCodeProfiles);
                PopulateWorkersCompRetentionProfile(segment, serverWorkersCompRetentionProfiles);

                PopulatePeriods(serverSegmentPair, segment);
                PopulateExposureSets(serverPackageId, segment, serverExposureSets);
                PopulateAggregateLossSets(serverPackageId, segment, serverAggregateLossSets);
                PopulateIndividualLossSets(serverPackageId, segment, serverIndividualLossSets);
                PopulateRateChangeSets(serverPackageId, segment, serverRateChangeSets);

                ExcelSheetActivateEventManager.RefreshRibbon(segment);
            }

            var summaryBuilder = new ProspectiveExposureSummaryBuilder();
            summaryBuilder.Build();

            package.SetAllItemsToNotDirty();
            package.IsSelected = true;
        }

        private static int CountStringItems(StringBuilder sb)
        {
            return Regex.Matches(sb.ToString(), Environment.NewLine).Count;
        }

        public void MarkItemsAsDone(StringBuilder sb)
        {
            sb.Replace("Rebuilding", "Rebuilt");
            sb.Replace(" ...", string.Empty);
        }

        public void RemoveFirstRow(StringBuilder sb)
        {
            sb.Remove(0, Convert.ToString(sb).Split('\n').First().Length + 1);
        }

        private static void PopulateProspectiveExposure(Segment segment, KeyValuePair<long, SubmissionSegmentModel> serverSegmentPair)
        {
            segment.ProspectiveExposureAmountExcelMatrix.GetInputRange().Value2 = serverSegmentPair.Value.SubjectBaseAmount;
        }

        private static void PopulateSublines(KeyValuePair<long, SubmissionSegmentModel> serverSegmentPair, ISegment segment)
        {
            var serverSublines = serverSegmentPair.Value.ValidSublines.ToList();
            var sublinesInRange = segment.SublineProfile.SublineExcelMatrix.GetBodyRange().GetContent().ForceContentToStrings()
                .GetColumn(0).ToList();

            var sublineCounter = 0;
            var sublineInRangeDictionary =
                sublinesInRange.ToDictionary(sublineInRange => sublineInRange, sublineInRange => sublineCounter++);
            var sublineReferenceData = SublineCodesFromBex.ReferenceData.ToDictionary(data => data.SublineId);

            var sublineValues = new object[sublinesInRange.Count, 1];
            foreach (var serverSubline in serverSublines)
            {
                var code = serverSubline.Id;
                var lobSublineName = SublineCodesFromBex.ConvertCodeToShortNameWithLob(sublineReferenceData[code].SublineId);

                var row = sublineInRangeDictionary[lobSublineName];
                var value = serverSubline.Value;

                sublineValues[row, 0] = value;
            }

            var excelMatrix = segment.SublineProfile.SublineExcelMatrix;
            excelMatrix.GetInputRange().Value2 = sublineValues;
            excelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;

        }

        private static void PopulateUmbrellaTypes(KeyValuePair<long, SubmissionSegmentModel> serverSegmentPair, ISegment segment)
        {
            if (!serverSegmentPair.Value.IsUmbrella) return;

            try
            {
                segment.IsCurrentlyRebuilding = true;
                segment.IsUmbrella = true;
            }
            finally
            {
                segment.IsCurrentlyRebuilding = false;
            }

            var referenceData = UmbrellaTypesFromBex.ReferenceData.ToDictionary(data => data.UmbrellaTypeCode);
            var serverUmbrellaTypes = serverSegmentPair.Value.UmbrellaTypeAllocations
                .Where(uta => !referenceData[uta.Id].IsPersonal)
                .OrderBy(uta => referenceData[uta.Id].DisplayOrder).ToList();
            if (!serverUmbrellaTypes.Any()) return;

            var excelMatrix = segment.UmbrellaExcelMatrix;
            var serverUmbrellaTypeCodes = serverUmbrellaTypes.Select(ut => ut.Id.ToString()).ToList();
            UmbrellaTypeAllocatorManager.ModifyWorksheetFromUmbrellaChanges(segment, serverUmbrellaTypeCodes);
               

            var values = new object[serverUmbrellaTypes.Count, 1];
            var row = 0;
            foreach (var serverUmbrellaType in serverUmbrellaTypes)
            {
                var value = serverUmbrellaType.Value;
                values[row++, 0] = value;
            }

            excelMatrix.GetInputRange().Value2 = values;
            excelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;
        }

        private IEnumerable<KeyValuePair<long, RateChangeSetModel>> CreateRateChangeSets(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverRateChangeSets = _bexCommunicationManager.GetServerRateChangeSets(serverPackageId, serverSegment.Id).OrderBy(spp => spp.Key).ToList();
            var rateChangeSetCounter = 0;

            foreach (var serverRateChangeSet in serverRateChangeSets)
            {
                var rateChangeSet = new RateChangeSet(segment.Id, rateChangeSetCounter)
                {
                    SourceId = serverRateChangeSet.Key, SourceTimestamp = DateTime.Now
                };
                rateChangeSet.ExcelMatrix.IntraDisplayOrder = rateChangeSetCounter++;

                var sublines = serverRateChangeSet.Value.Sublines.Select(id => _allSublines[Convert.ToInt32(id)]).ToList();
                rateChangeSet.ExcelMatrix.UpdateSublines(sublines);
                segment.ExcelComponents.Add(rateChangeSet);
            }

            //rate change sets may be blank and therefore not on the server
            //put all the missing rate changes in a blank set
            var serverRateChangeSetsSublines = serverRateChangeSets.SelectMany(rcs => rcs.Value.Sublines);
            var serverSegmentSublines = serverSegment.ValidSublines.Select(ss => ss.Id);
            var missingSublineIds = serverSegmentSublines.Except(serverRateChangeSetsSublines);
            var missingSublines = missingSublineIds.Select(id => _allSublines[Convert.ToInt32(id)]).ToList();
            if (missingSublines.Any())
            {
                var rateChangeSet = new RateChangeSet(segment.Id, rateChangeSetCounter);
                rateChangeSet.ExcelMatrix.IntraDisplayOrder = rateChangeSetCounter;
                rateChangeSet.ExcelMatrix.UpdateSublines(missingSublines);
                segment.ExcelComponents.Add(rateChangeSet);
            }

            return serverRateChangeSets;
        }

        private IEnumerable<KeyValuePair<long, IndividualLossSetModel>> CreateIndividualLossSets(
            long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverLossSets = _bexCommunicationManager.GetServerIndividualLossSets(serverPackageId, serverSegment.Id)
                .OrderBy(spp => spp.Key).ToList();
            var lossSetCounter = 0;

            if (serverLossSets.Any())
            {
                foreach (var serverLossSet in serverLossSets)
                {
                    var lossSet = new IndividualLossSet(segment.Id, lossSetCounter)
                    {
                        SourceId = serverLossSet.Key, SourceTimestamp = DateTime.Now
                    };
                    lossSet.ExcelMatrix.IntraDisplayOrder = lossSetCounter++;

                    var lossSetSublines = serverLossSet.Value.Sublines.Select(id => _allSublines[Convert.ToInt32(id)]).ToList();
                    lossSet.ExcelMatrix.UpdateSublines(lossSetSublines);
                    segment.ExcelComponents.Add(lossSet);
                }
            }
            else
            {
                //add one blank individual loss set if necessary
                var sublines = segment.Select(id => _allSublines[Convert.ToInt32(id.Code)]).ToList();
                var lossSet = new IndividualLossSet(segment.Id, lossSetCounter);
                lossSet.ExcelMatrix.UpdateSublines(sublines);
                segment.ExcelComponents.Add(lossSet);
            }

            return serverLossSets;
        }

        private IEnumerable<KeyValuePair<long, AggregateLossSetModel>> CreateAggregateLossSets(
            long serverPackageId,
            SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverLossSets = _bexCommunicationManager.GetServerAggregateLossSets(serverPackageId, serverSegment.Id)
                .OrderBy(spp => spp.Key).ToList();
            var lossSetCounter = 0;

            if (serverLossSets.Any())
            {
                foreach (var serverLossSet in serverLossSets)
                {
                    var lossSet = new AggregateLossSet(segment.Id, lossSetCounter)
                    {
                        SourceId = serverLossSet.Key, SourceTimestamp = DateTime.Now
                    };
                    lossSet.ExcelMatrix.IntraDisplayOrder = lossSetCounter++;

                    var lossSetSublines =
                        serverLossSet.Value.Sublines.Select(id => _allSublines[Convert.ToInt32(id)]).ToList();
                    lossSet.ExcelMatrix.UpdateSublines(lossSetSublines);
                    segment.ExcelComponents.Add(lossSet);
                }
            }
            else
            {
                //add one blank aggregate loss set if necessary
                var lossSetSublines = segment.Select(id => _allSublines[Convert.ToInt32(id.Code)]).ToList();
                var lossSet = new AggregateLossSet(segment.Id, lossSetCounter);
                lossSet.ExcelMatrix.UpdateSublines(lossSetSublines);
                segment.ExcelComponents.Add(lossSet);
            }

            return serverLossSets;
        }

        private IEnumerable<KeyValuePair<long, ExposureSetModel>> CreateExposureSets(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverExposureSets = _bexCommunicationManager.GetServerExposureSets(serverPackageId, serverSegment.Id).OrderBy(spp => spp.Key)
                .ToList();
            var exposureSetCounter = 0;

            if (serverExposureSets.Any())
            {
                foreach (var serverExposureSet in serverExposureSets)
                {
                    var exposureSet = new ExposureSet(segment.Id, exposureSetCounter)
                    {
                        SourceId = serverExposureSet.Key, SourceTimestamp = DateTime.Now
                    };
                    exposureSet.ExcelMatrix.IntraDisplayOrder = exposureSetCounter++;

                    var exposureSetSublines = serverExposureSet.Value.Sublines.Select(id => _allSublines[Convert.ToInt32(id)]).ToList();
                    exposureSet.ExcelMatrix.UpdateSublines(exposureSetSublines);
                    segment.ExcelComponents.Add(exposureSet);
                }
            }
            else
            {
                //add one blank exposure set if necessary
                var sublines = segment.Select(id => _allSublines[Convert.ToInt32(id.Code)]).ToList();
                var exposureSet = new ExposureSet(segment.Id, exposureSetCounter);
                exposureSet.ExcelMatrix.UpdateSublines(sublines);
                segment.ExcelComponents.Add(exposureSet);
            }

            return serverExposureSets;
        }

        private IEnumerable<KeyValuePair<long, StateDistributionModel>> CreateStateProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager.GetServerStateProfiles(serverPackageId, serverSegment.Id).OrderBy(spp => spp.Key)
                .ToList();
            var profileCounter = 0;

            if (serverProfiles.Any())
            {
                foreach (var serverProfile in serverProfiles)
                {
                    var profile = new StateProfile(segment.Id, profileCounter)
                    {
                        SourceId = serverProfile.Key, SourceTimestamp = DateTime.Now
                    };
                    profile.ExcelMatrix.IntraDisplayOrder = profileCounter++;

                    var profileSublines =
                        serverProfile.Value.SublineIds.Select(id => _allSublines[Convert.ToInt32(id)]).ToList();
                    profile.ExcelMatrix.UpdateSublines(profileSublines);
                    segment.ExcelComponents.Add(profile);
                }
            }
            else
            {
                //add one blank state profile if necessary
                var sublines = segment.Select(id => _allSublines[Convert.ToInt32(id.Code)]).ToList();
                if (sublines.Any(sl => sl.HasStateProfile))
                {
                    var profile = new StateProfile(segment.Id, profileCounter);
                    profile.ExcelMatrix.UpdateSublines(sublines.Where(sl => sl.HasStateProfile).ToList());
                    segment.ExcelComponents.Add(profile);
                }
            }

            return serverProfiles;
        }


        private IEnumerable<KeyValuePair<long, ConstructionTypeDistributionModel>> CreateConstructionTypeProfiles(long serverPackageId,
            SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager
                .GetServerConstructionTypeProfiles(serverPackageId, serverSegment.Id)
                .OrderBy(serverProfile => serverProfile.Key)
                .ToList();


            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.ConstructionTypeProfiles.First(hp => hp.ComponentId == serverProfile.Value.SublineIds.First());
                profile.SourceId = serverProfile.Key;
                profile.SourceTimestamp = DateTime.Now;
            }
        
            return serverProfiles;
        }

        private IEnumerable<KeyValuePair<long, OccupancyTypeDistributionModel>> CreateOccupancyTypeProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager
                .GetServerOccupancyTypeProfiles(serverPackageId, serverSegment.Id)
                .OrderBy(serverProfile => serverProfile.Key)
                .ToList();
            
                foreach (var serverProfile in serverProfiles)
                {
                    var profile = segment.OccupancyTypeProfiles.First(prof => prof.ComponentId == serverProfile.Value.SublineIds.First());
                    profile.SourceId = serverProfile.Key;
                    profile.SourceTimestamp = DateTime.Now;
                }
            
            return serverProfiles;
        }

        private IEnumerable<KeyValuePair<long, ProtectionClassDistributionModel>> CreateProtectionClassProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager
                .GetServerProtectionClassProfiles(serverPackageId, serverSegment.Id)
                .OrderBy(serverProfile => serverProfile.Key)
                .ToList();
            
            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.ProtectionClassProfiles.First(hp => hp.ComponentId == serverProfile.Value.SublineIds.First());
                profile.SourceId = serverProfile.Key;
                profile.SourceTimestamp = DateTime.Now;
            }
            
            return serverProfiles;
        }

        private IEnumerable<KeyValuePair<long, TotalInsuredValueDistributionModel>> CreateTotalInsuredValueProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager
                .GetServerTotalInsuredValueProfiles(serverPackageId, serverSegment.Id)
                .OrderBy(serverProfile => serverProfile.Key)
                .ToList();

            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.TotalInsuredValueProfiles.First(insuredValueProfile => insuredValueProfile.ComponentId == serverProfile.Value.SublineIds.First());
                profile.SourceId = serverProfile.Key;
                profile.SourceTimestamp = DateTime.Now;
            }

            return serverProfiles;
        }

        private IEnumerable<KeyValuePair<long, WorkersCompStateHazardGroupDistributionModel>> CreateWorkersCompStateHazardGroupProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager
                .GetServerWorkersCompStateHazardGroupProfiles(serverPackageId, serverSegment.Id)
                .OrderBy(serverProfile => serverProfile.Key)
                .ToList();

            //at most one
            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.WorkersCompStateHazardGroupProfile;
                profile.SourceId = serverProfile.Key;
                profile.SourceTimestamp = DateTime.Now;
                profile.ExcelMatrix.IsIndependent = false;
            }

            return serverProfiles;
        }

        private IEnumerable<KeyValuePair<long, WorkersCompStateClassCodeDistributionModel>> CreateWorkersCompClassCodeProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager
                .GetServerWorkersCompClassCodeProfiles(serverPackageId, serverSegment.Id)
                .OrderBy(serverProfile => serverProfile.Key)
                .ToList();

            //at most one
            foreach (var serverProfile in serverProfiles)
            {
                if (segment.WorkersCompClassCodeProfile == null)
                {
                    segment.ExcelComponents.Add(new WorkersCompClassCodeProfile(segment.Id));
                }

                var profile = segment.WorkersCompClassCodeProfile;
                Debug.Assert(profile != null, nameof(profile) + " != null");
                profile.SourceId = serverProfile.Key;
                profile.SourceTimestamp = DateTime.Now;
            }

            return serverProfiles;
        }

        private IEnumerable<KeyValuePair<long, WorkersCompStateAttachmentDistributionModel>> CreateWorkersCompRetentionProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager
                .GetServerWorkersCompStateAttachmentProfiles(serverPackageId, serverSegment.Id)
                .OrderBy(serverProfile => serverProfile.Key)
                .ToList();

            //at most one
            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.WorkersCompStateAttachmentProfile;
                profile.SourceId = serverProfile.Key;
                profile.SourceTimestamp = DateTime.Now;
            }

            return serverProfiles;
        }
        
        private IEnumerable<KeyValuePair<long, HazardDistributionModel>> CreateHazardProfiles(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles = _bexCommunicationManager.GetServerHazardProfiles(serverPackageId, serverSegment.Id).OrderBy(spp => spp.Key)
                .ToList();
            
            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.HazardProfiles.First(hp => hp.ComponentId == serverProfile.Value.SublineIds.First());
                profile.SourceId = serverProfile.Key;
                profile.SourceTimestamp = DateTime.Now;
            }

            return serverProfiles;
        }

        private IEnumerable<KeyValuePair<long, PolicyDistributionModel>> CreatePolicyProfiles(long serverPackageId,
            SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var serverProfiles =
                _bexCommunicationManager.GetServerPolicyProfiles(serverPackageId, serverSegment.Id).OrderBy(spp => spp.Key).ToList();
            var profileCounter = 0;

            if (serverProfiles.Any())
            {
                foreach (var serverProfile in serverProfiles)
                {
                    var profile = new PolicyProfile(segment.Id, profileCounter)
                    {
                        SourceId = serverProfile.Key,
                        SourceTimestamp = DateTime.Now
                    };
                    profile.ExcelMatrix.IntraDisplayOrder = profileCounter++;

                    if (serverProfile.Value.UmbrellaTypeId.HasValue)
                    {
                        profile.UmbrellaType = Convert.ToInt32(serverProfile.Value.UmbrellaTypeId.Value);
                        var umbrellaTypeName = UmbrellaTypesFromBex.GetName(profile.UmbrellaType.Value);
                        profile.Name = $"{BexConstants.PolicyProfileName} " +
                                       $"{UmbrellaTypesFromBex.AbbreviateUmbrellaType(umbrellaTypeName)}";
                    }

                    var profileSublines =
                        serverProfile.Value.SublineIds.Select(id => _allSublines[Convert.ToInt32(id)]).ToList();
                    profile.ExcelMatrix.UpdateSublines(profileSublines);
                    if (serverProfile.Value.UmbrellaTypeId.HasValue)
                        profile.UmbrellaType = Convert.ToInt32(serverProfile.Value.UmbrellaTypeId);
                    segment.ExcelComponents.Add(profile);
                }
            }
            else
            {
                //add one blank policy profile if necessary
                var sublines = segment.Select(id => _allSublines[Convert.ToInt32(id.Code)]).ToList();
                if (sublines.Any(sl => sl.HasPolicyProfile))
                {
                    var profile = new PolicyProfile(segment.Id, profileCounter);
                    profile.ExcelMatrix.UpdateSublines(sublines);
                    segment.ExcelComponents.Add(profile);
                }
            }

            return serverProfiles;
        }

        private void PopulateIndividualLossSets(long serverPackageId, ISegment segment, IEnumerable<KeyValuePair<long, IndividualLossSetModel>> serverIndividualLossSets)
        {
            var lossSetCounter = 0;
            foreach (var serverLossSet in serverIndividualLossSets)
            {
                var serverLosses = _bexCommunicationManager.GetServerIndividualLosses(serverPackageId, segment.SourceId.Value, serverLossSet.Key).Values;
                var lossSetDescriptor = segment.IndividualLossSetDescriptor;

                var columnCounter = 0;
                var occurrenceColumn = columnCounter++;
                var claimColumn = columnCounter++;
                var eventCodeColumn = lossSetDescriptor.IsEventCodeAvailable ? columnCounter++ : new int?();
                var descriptionColumn = columnCounter++;
                var accidentDateColumn = lossSetDescriptor.IsAccidentDateAvailable ? columnCounter++ : new int?();
                var policyDateColumn = lossSetDescriptor.IsPolicyDateAvailable ? columnCounter++ : new int?();
                var reportDateColumn = lossSetDescriptor.IsReportDateAvailable ? columnCounter++ : new int?();
                var limitColumn = lossSetDescriptor.IsPolicyLimitAvailable ? columnCounter++ : new int?();
                var attachmentColumn = lossSetDescriptor.IsPolicyAttachmentAvailable ? columnCounter++ : new int?();
                var paidLossColumn = lossSetDescriptor.IsPaidAvailable && !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var paidAlaeColumn = lossSetDescriptor.IsPaidAvailable && !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var paidLossAndAlaeColumn = lossSetDescriptor.IsPaidAvailable && lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var reportedLossColumn = !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var reportedAlaeColumn = !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var reportedLossAndAlaeColumn = lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();

                var losses = new object[serverLosses.Count, columnCounter];

                var lossSet = segment.IndividualLossSets.First(als => als.ComponentId == lossSetCounter); 
                var lossCounter = 0;

                var customOrderBy = GetCustomOrderBy(segment);
                foreach (var serverLoss in serverLosses.OrderBy(customOrderBy))
                {
                    losses[lossCounter, occurrenceColumn] = serverLoss.OccurrenceId;
                    losses[lossCounter, claimColumn] = serverLoss.ClaimNumber;
                    if (eventCodeColumn.HasValue) losses[lossCounter, eventCodeColumn.Value] = serverLoss.EventCode;
                    losses[lossCounter, descriptionColumn] = serverLoss.LossDescription;
                    if (accidentDateColumn.HasValue) losses[lossCounter, accidentDateColumn.Value] = serverLoss.AccidentDate?.Date;
                    if (policyDateColumn.HasValue) losses[lossCounter, policyDateColumn.Value] = serverLoss.PolicyDate?.Date;
                    if (reportDateColumn.HasValue) losses[lossCounter, reportDateColumn.Value] = serverLoss.ReportedDate?.Date;
                    if (limitColumn.HasValue) losses[lossCounter, limitColumn.Value] = serverLoss.PolicyLimitAmount;
                    if (attachmentColumn.HasValue) losses[lossCounter, attachmentColumn.Value] = serverLoss.PolicyAttachmentAmount;
                    
                    if (paidLossColumn.HasValue) losses[lossCounter, paidLossColumn.Value] = serverLoss.PaidLossAmount;
                    if (paidAlaeColumn.HasValue) losses[lossCounter, paidAlaeColumn.Value] = serverLoss.PaidAlaeAmount;
                    if (paidLossAndAlaeColumn.HasValue) losses[lossCounter, paidLossAndAlaeColumn.Value] = serverLoss.PaidCombinedAmount;
                    if (reportedLossColumn.HasValue) losses[lossCounter, reportedLossColumn.Value] = serverLoss.ReportedLossAmount;
                    if (reportedAlaeColumn.HasValue) losses[lossCounter, reportedAlaeColumn.Value] = serverLoss.ReportedAlaeAmount;
                    if (reportedLossAndAlaeColumn.HasValue) losses[lossCounter, reportedLossAndAlaeColumn.Value] = serverLoss.ReportedCombinedAmount;

                    lossSet.Ledger.Add(new LedgerItem { RowId = lossCounter, SourceId = serverLoss.Id, SourceTimestamp = DateTime.Now });
                    lossCounter++;
                }

                
                var inputRange = lossSet.ExcelMatrix.GetInputRange();
                if (inputRange.Rows.Count < serverLosses.Count)
                {
                    var rowCount = serverLosses.Count - inputRange.Rows.Count + 1;
                    var rangeInserter = new IndividualLossSetRowInserter
                    {
                        ExcelMatrix = lossSet.ExcelMatrix,
                        ExcelRange = inputRange,
                        RowCount = rowCount,
                        StartRow = 3,
                        IsSelectionOnSecondRow = false,
                        IgnoreLedger = true
                    };
                    using (new ExcelEventDisabler())
                    {
                        rangeInserter.ModifyRange();
                    }
                }

                for (var i = serverLosses.Count; i < inputRange.Rows.Count; i++)
                {
                    lossSet.Ledger.Add(new LedgerItem { RowId = i });
                }
                inputRange.Resize[lossCounter, columnCounter].Value2 = losses;
                lossSet.Ledger.SetToNotDirty();

                if (serverLossSet.Value.Threshold.HasValue)
                {
                    var thresholdRange = ExcelMatrix.GetThresholdRangeName(segment.Id, lossSet.ComponentId);
                    thresholdRange.GetTopRightCell().Value2 = serverLossSet.Value.Threshold;
                }

                lossSetCounter++;
            }
        }

        private void PopulateRateChangeSets(long serverPackageId, ISegment segment, IEnumerable<KeyValuePair<long, RateChangeSetModel>> serverRateChangeSets )
        {
            var rateChangeSetCounter = 0;
            foreach (var serverRateChangeSet in serverRateChangeSets)
            {
                Debug.Assert(segment.SourceId != null, "segment.SourceId != null");
                var serverRateChanges = _bexCommunicationManager.GetServerRateChanges(serverPackageId, segment.SourceId.Value, serverRateChangeSet.Key).Values;
                var rateChanges = new object[serverRateChanges.Count, 2];

                var rateChangeSet = segment.RateChangeSets.First(rcs => rcs.ComponentId == rateChangeSetCounter); 
                var rateChangeCounter = 0;
                foreach (var serverRateChange in serverRateChanges.OrderBy(se => se.EffectiveDate))
                {
                    rateChanges[rateChangeCounter, 0] = serverRateChange.EffectiveDate.Date;
                    rateChanges[rateChangeCounter, 1] = serverRateChange.Value;
                    rateChangeSet.Ledger.Add(new LedgerItem { RowId = rateChangeCounter, IsDirty = false, SourceId = serverRateChange.Id, SourceTimestamp = DateTime.Now });
                    rateChangeCounter++;
                }

                var inputRange = rateChangeSet.ExcelMatrix.GetInputRange();
                if (inputRange.Rows.Count < serverRateChanges.Count)
                {
                    var rowCount = serverRateChanges.Count - inputRange.Rows.Count;
                    var rangeInserter = new RateChangeSetRowInserter
                    {
                        ExcelMatrix = rateChangeSet.ExcelMatrix,
                        ExcelRange = inputRange,
                        RowCount = rowCount,
                        StartRow = 3,
                        IsSelectionOnSecondRow = false,
                        IgnoreLedger = true
                    };
                    using (new ExcelEventDisabler())
                    {
                        rangeInserter.ModifyRange();
                    }
                }

                for (var i = serverRateChanges.Count; i < inputRange.Rows.Count; i++)
                {
                    rateChangeSet.Ledger.Add(new LedgerItem { RowId = i });
                }
                inputRange.Resize[rateChangeCounter, 2].Value2 = rateChanges;
                rateChangeSet.Ledger.SetToNotDirty();
                rateChangeSetCounter++;
            }
        }

        private void PopulateAggregateLossSets(long serverPackageId, ISegment segment, IEnumerable<KeyValuePair<long, AggregateLossSetModel>> serverAggregateLossSets)
        {
            var lossSetCounter = 0;
            foreach (var serverLossSet in serverAggregateLossSets)
            {
                Debug.Assert(segment.SourceId != null, "segment.SourceId != null");
                var serverLosses = _bexCommunicationManager.GetServerAggregateLosses(serverPackageId, segment.SourceId.Value, serverLossSet.Key).Values;
                var lossSet = segment.AggregateLossSets.First(als => als.ComponentId == lossSetCounter);

                var lossSetDescriptor = segment.AggregateLossSetDescriptor;

                var columnCounter = 0;
                var paidLossColumn = lossSetDescriptor.IsPaidAvailable && !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var paidAlaeColumn = lossSetDescriptor.IsPaidAvailable && !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var paidLossAndAlaeColumn = lossSetDescriptor.IsPaidAvailable && lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var reportedLossColumn = !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var reportedAlaeColumn = !lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();
                var reportedLossAndAlaeColumn = lossSetDescriptor.IsLossAndAlaeCombined ? columnCounter++ : new int?();

                var losses = new object[serverLosses.Count, columnCounter];
                var lossCounter = 0;
                foreach (var serveLoss in serverLosses.OrderBy(se => se.StartDate))
                {
                    if (paidLossColumn.HasValue) losses[lossCounter, paidLossColumn.Value] = serveLoss.PaidLossAmount;
                    if (paidAlaeColumn.HasValue) losses[lossCounter, paidAlaeColumn.Value] = serveLoss.PaidAlaeAmount;
                    if (paidLossAndAlaeColumn.HasValue) losses[lossCounter, paidLossAndAlaeColumn.Value] = serveLoss.PaidCombinedAmount;
                    if (reportedLossColumn.HasValue) losses[lossCounter, reportedLossColumn.Value] = serveLoss.ReportedLossAmount;
                    if (reportedAlaeColumn.HasValue) losses[lossCounter, reportedAlaeColumn.Value] = serveLoss.ReportedAlaeAmount;
                    if (reportedLossAndAlaeColumn.HasValue)                        losses[lossCounter, reportedLossAndAlaeColumn.Value] = serveLoss.ReportedCombinedAmount;

                    lossSet.Ledger.Add(new LedgerItem { RowId = lossCounter, IsDirty = false, SourceId = serveLoss.Id, SourceTimestamp = DateTime.Now });
                    lossCounter++;
                }

                var inputRange = lossSet.ExcelMatrix.GetInputRange();
                for (var i = serverLosses.Count; i < inputRange.Rows.Count; i++)
                {
                    lossSet.Ledger.Add(new LedgerItem { RowId = i });
                }
                inputRange.Resize[lossCounter, columnCounter].Value2 = losses;
                lossSet.Ledger.SetToNotDirty();
                lossSetCounter++;
            }
        }

        private void PopulateExposureSets(long serverPackageId, ISegment segment, IEnumerable<KeyValuePair<long, ExposureSetModel>> serverExposureSets)
        {
            var exposureSetCounter = 0;
            foreach (var serverExposureSet in serverExposureSets)
            {
                Debug.Assert(segment.SourceId != null, "segment.SourceId != null");
                var serverExposures = _bexCommunicationManager.GetServerExposures(serverPackageId, segment.SourceId.Value, serverExposureSet.Key).Values;
                var exposures = new object[serverExposures.Count, 1];

                var exposureSet = segment.ExposureSets.First(pp => pp.ComponentId == exposureSetCounter);
                var exposureCounter = 0;
                foreach (var serverExposure in serverExposures.OrderBy(se => se.StartDate))
                {
                    exposures[exposureCounter, 0] = serverExposure.Amount;
                    exposureSet.Ledger.Add(new LedgerItem { RowId = exposureCounter, SourceId = serverExposure.Id, SourceTimestamp = DateTime.Now });
                    exposureCounter++;
                }

                var inputRange = exposureSet.ExcelMatrix.GetInputRange();
                for (var i = serverExposures.Count; i < inputRange.Rows.Count; i++)
                {
                    exposureSet.Ledger.Add(new LedgerItem{RowId = i});
                }
                inputRange.Resize[serverExposures.Count, 1].Value2 = exposures;
                exposureSet.Ledger.SetToNotDirty();
                exposureSetCounter++;
            }
        }

        private static void PopulatePolicyProfiles(ISegment segment, IEnumerable<KeyValuePair<long, PolicyDistributionModel>> serverPolicyProfiles)
        {
            var profileCounter = 0;
            foreach (var serverProfile in serverPolicyProfiles)
            {
                var items = serverProfile.Value.Items;
                var policies = new object[items.Count, 3];
                var policyCount = 0;
                foreach (var item in items)
                {
                    policies[policyCount, 0] = item.Limit;
                    policies[policyCount, 1] = item.Attachment;
                    policies[policyCount++, 2] = item.Value;
                }

                var profile = segment.PolicyProfiles.First(pp => pp.ComponentId == profileCounter);
                var inputRange = profile.ExcelMatrix.GetInputRange();
                if (inputRange.Rows.Count < items.Count)
                {
                    var rowCount = items.Count - inputRange.Rows.Count + 1;
                    var rangeInserter = new PolicyProfileRowInserter
                    {
                        ExcelRange = inputRange,
                        RowCount = rowCount,
                        StartRow = 3,
                        IsSelectionOnSecondRow = false
                    };
                    using (new ExcelEventDisabler())
                    {
                        rangeInserter.ModifyRange();
                    }
                }

                inputRange.Resize[items.Count, 3].Value2 = policies;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;
                profileCounter++;
            }
        }

        private static void PopulatePeriods(KeyValuePair<long, SubmissionSegmentModel> serverSegmentPair, Segment segment)
        {
            var serverPeriods = serverSegmentPair.Value.ValidPeriods.OrderBy(period => period.StartDate).ToList();

            if (!serverPeriods.Any()) return;

            var periodCounter = 0;
            var values = new object[serverPeriods.Count, 3];
            foreach (var serverPeriod in serverPeriods)
            {
                values[periodCounter, 0] = serverPeriod.StartDate.Date;
                values[periodCounter, 1] = serverPeriod.EndDate.Date;
                values[periodCounter, 2] = serverPeriod.EvaluationDate.Date;
                periodCounter++;
            }

            var inputRange = segment.PeriodSet.ExcelMatrix.GetInputRange();
            if (inputRange.Rows.Count < serverPeriods.Count)
            {
                var rowCount = serverPeriods.Count - inputRange.Rows.Count;
                var rangeInserter = new PeriodSetRowInserter
                {
                    ExcelMatrix = segment.PeriodSet.ExcelMatrix,
                    ExcelRange = inputRange,
                    RowCount = rowCount,
                    StartRow = 3,
                    IsSelectionOnSecondRow = false
                };
                using (new ExcelEventDisabler())
                {
                    rangeInserter.ModifyRange();
                }
            }

            inputRange.Resize[serverPeriods.Count, 3].Value2 = values;
        }

        
        private static void PopulateHazardProfiles(ISegment segment, IEnumerable<KeyValuePair<long, HazardDistributionModel>> serverHazardProfiles)
        {
            foreach (var serverProfile in serverHazardProfiles)
            {
                var profile = segment.HazardProfiles.First(pp => pp.ComponentId == serverProfile.Value.SublineIds.First());
                var hazardsInRange = profile.ExcelMatrix.GetBodyRange().GetContent().ForceContentToStrings().GetColumn(0).ToList();

                var hazardCounter = 0;
                var hazardsInRangeDictionary = hazardsInRange.ToDictionary(hazardInRange => hazardInRange, hazardInRange => hazardCounter++);
                var referenceData = HazardCodesFromBex.ReferenceData.Where(data => data.SubLineOfBusinessCode == profile.ComponentId).ToDictionary(data => data.Id);

                var items = serverProfile.Value.Items;
                var values = new object[hazardsInRange.Count, 1];
                foreach (var item in items)
                {
                    var code = item.HazardId;
                    var hazardName = referenceData[code].Name;

                    var row = hazardsInRangeDictionary[hazardName];
                    var value = item.Value;

                    values[row, 0] = value;
                }

                profile.ExcelMatrix.GetInputRange().Value2 = values;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;

            }
        }

        private static void PopulateStateProfiles(ISegment segment, IEnumerable<KeyValuePair<long, StateDistributionModel>> serverStateProfiles)
        {
            var referenceData = StateCodesFromBex.ReferenceData.ToDictionary(data => data.Id);
            var profileCounter = 0;
            foreach (var serverProfile in serverStateProfiles)
            {
                var profile = segment.StateProfiles.First(pp => pp.ComponentId == profileCounter);
                var statesInRange = profile.ExcelMatrix.GetBodyRange().GetColumn(0).GetContent().ForceContentToStrings()
                    .GetColumn(0).ToList();

                var stateCounter = 0;
                var statesInRangeDictionary = statesInRange.ToDictionary(stateInRange => stateInRange, stateInRange => stateCounter++);

                var items = serverProfile.Value.Items;
                var values = new object[statesInRange.Count, 1];
                foreach (var item in items)
                {
                    var code = item.StateCode;
                    var abbreviation = referenceData[code].Abbreviation;

                    var row = statesInRangeDictionary[abbreviation];
                    var value = item.Value;

                    values[row, 0] = value;
                }

                profile.ExcelMatrix.GetInputRange().Value2 = values;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;
                profileCounter++;
            }
        }


        private static void PopulateConstructionTypeProfiles(ISegment segment, IEnumerable<KeyValuePair<long, ConstructionTypeDistributionModel>> serverConstructionTypeProfiles)
        {
            foreach (var serverProfile in serverConstructionTypeProfiles)
            {
                var profile = segment.ConstructionTypeProfiles.First(pp => pp.ComponentId == serverProfile.Value.SublineIds.First());
                var valuesInRange = profile.ExcelMatrix.GetBodyRange().GetContent().ForceContentToStrings().GetColumn(0).ToList();

                var counter = 0;
                var rangeDictionary = valuesInRange.ToDictionary(valueInRange => valueInRange, valueInRange => counter++);
                var referenceData = ConstructionTypeCodesFromBex.ReferenceData.Where(data => data.SubLineOfBusinessCode == profile.ComponentId).ToDictionary(data => data.Id);

                var items = serverProfile.Value.Items;
                var values = new object[valuesInRange.Count, 1];
                foreach (var item in items)
                {
                    var code = item.ConstructionTypeId;
                    var name = referenceData[code].Name;

                    var row = rangeDictionary[name];
                    var value = item.Weight;

                    values[row, 0] = value;
                }

                profile.ExcelMatrix.GetInputRange().Value2 = values;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;
            }
        }

        private static void PopulateOccupancyTypeProfiles(ISegment segment, IEnumerable<KeyValuePair<long, OccupancyTypeDistributionModel>> serverOccupancyTypeProfiles)
        {
            var profileCounter = 0;
            foreach (var serverProfile in serverOccupancyTypeProfiles)
            {
                var profile = segment.OccupancyTypeProfiles.First(pp => pp.ComponentId == profileCounter);
                var valuesInRange = profile.ExcelMatrix.GetBodyRange().GetContent().ForceContentToStrings().GetColumn(0)
                    .ToList();

                var counter = 0;
                var rangeDictionary = valuesInRange.ToDictionary(valueInRange => valueInRange, valueInRange => counter++);
                var referenceData = OccupancyTypeCodesFromBex.ReferenceData.ToDictionary(data => data.Id);

                var items = serverProfile.Value.Items;
                var values = new object[valuesInRange.Count, 1];
                foreach (var item in items)
                {
                    var code = item.OccupancyTypeId;
                    var name = referenceData[code].Name;

                    var row = rangeDictionary[name];
                    var value = item.Weight;

                    values[row, 0] = value;
                }

                profile.ExcelMatrix.GetInputRange().Value2 = values;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;
                profileCounter++;
            }
        }

        private static void PopulateProtectionClassProfiles(ISegment segment, IEnumerable<KeyValuePair<long, ProtectionClassDistributionModel>> serverProtectionClassProfiles)
        {
            foreach (var serverProfile in serverProtectionClassProfiles)
            {
                var profile = segment.ProtectionClassProfiles.First(pp => pp.ComponentId == serverProfile.Value.SublineIds.First());
                var valuesInRange = profile.ExcelMatrix.GetBodyRange().GetContent().ForceContentToStrings().GetColumn(0).ToList();

                var counter = 0;
                var rangeDictionary = valuesInRange.ToDictionary(valueInRange => valueInRange, valueInRange => counter++);
                var referenceData = ProtectionClassCodesFromBex.ReferenceData.Where(data => data.SubLineOfBusinessCode == profile.ComponentId).ToDictionary(data => data.Id);

                var items = serverProfile.Value.Items;
                var values = new object[valuesInRange.Count, 1];
                foreach (var item in items)
                {
                    var code = item.ProtectionClassId;
                    var name = referenceData[code].Name;

                    var row = rangeDictionary[name];
                    var value = item.Weight;

                    values[row, 0] = value;
                }

                profile.ExcelMatrix.GetInputRange().Value2 = values;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;
            }
        }

        private static void PopulateTotalInsuredValueProfiles(ISegment segment, IEnumerable<KeyValuePair<long, TotalInsuredValueDistributionModel>> serverProfiles)
        {
            if (!segment.ContainsProperty()) return;

            var profileCounter = 0;
            foreach (var serverProfile in serverProfiles)
            {
                var items = serverProfile.Value.Items;
                var totalInsuredValueContent = new object[items.Count, 5];
                var tivRowCount = 0;
                foreach (var item in items)
                {
                    totalInsuredValueContent[tivRowCount, 0] = item.TotalInsuredValue;
                    totalInsuredValueContent[tivRowCount, 1] = item.Share;
                    totalInsuredValueContent[tivRowCount, 2] = item.Limit;
                    totalInsuredValueContent[tivRowCount, 3] = item.Attachment;
                    totalInsuredValueContent[tivRowCount, 4] = item.Weight;
                    tivRowCount++;
                }
                
                var profile = segment.TotalInsuredValueProfiles.First(pp => pp.ComponentId == profileCounter);
                profile.IsExpanded = true;
                profile.ExcelMatrix.SynchronizeExpansion();

                var inputRange = profile.ExcelMatrix.GetInputRange();
                if (inputRange.Rows.Count < items.Count)
                {
                    var rowCount = items.Count - inputRange.Rows.Count + 1;
                    var rangeInserter = new TotalInsuredValueProfileRowInserter
                    {
                        ExcelRange = inputRange,
                        RowCount = rowCount,
                        StartRow = 3,
                        IsSelectionOnSecondRow = false
                    };
                    using (new ExcelEventDisabler())
                    {
                        rangeInserter.ModifyRange();
                    }
                }

                inputRange.Resize[items.Count, 5].Value2 = totalInsuredValueContent;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentName;
                profileCounter++;
            }
        }

        private static void PopulateMinnesotaRetention(MinnesotaRetention minnesotaRetention, long? id)
        {
            if (!id.HasValue) return;

            var match = MinnesotaRetentionsFromBex.ReferenceData.SingleOrDefault(data => data.Id == id);
            if (match != null)
            {
                minnesotaRetention.ExcelMatrix.GetInputRange().GetTopLeftCell().Value = match.RetentionAmount;
            }
        }

        private static void PopulateWorkersCompStateHazardGroupProfile(ISegment segment, IEnumerable<KeyValuePair<long, WorkersCompStateHazardGroupDistributionModel>> serverProfiles)
        {
            //at most one
            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.WorkersCompStateHazardGroupProfile;
                var states = StateCodesFromBex.GetWorkersCompStates().OrderBy(state => state.DisplayOrder).ToList();
                var hazardGroups = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.OrderBy(hg => hg.DisplayOrder).ToList();

                var stateCounter = 0;
                var stateDictionary = states.ToDictionary(key => Convert.ToInt64(key.Id), value => stateCounter++);

                var hazardGroupCounter = 0;
                var hazardGroupDictionary = hazardGroups.ToDictionary(key => Convert.ToInt64(key.Id), value => hazardGroupCounter++);

                var items = serverProfile.Value.Items;
                var values = new object[states.Count, hazardGroups.Count];
                foreach (var item in items)
                {
                    var stateId = item.WorkersCompStateId;
                    var hazardGroupId = item.WorkersCompHazardGroupId;

                    var row = stateDictionary[stateId];
                    var column = hazardGroupDictionary[hazardGroupId];
                    var value = item.Value;

                    values[row, column] = value;
                }

                profile.ExcelMatrix.GetInputRange().Value2 = values;
                profile.ExcelMatrix.BasisRangeName.GetRange().Value = ProfileBasisFromBex.ReferenceData.Single(item => item.Id == BexConstants.PremiumProfileBasisId).Name;
            }
        }

        //todo wc rebuild
        private static void PopulateWorkersCompClassCodeProfile(ISegment segment, IEnumerable<KeyValuePair<long, WorkersCompStateClassCodeDistributionModel>> serverProfiles)
        {
            var serverProfilesAsList = serverProfiles.ToList();
            if (!serverProfilesAsList.Any()) return;

            var states = StateCodesFromBex.GetWorkersCompStates().OrderBy(state => state.DisplayOrder).ToList();
            var stateAbbreviations = states.ToDictionary(key => Convert.ToInt64(key.Id), value => value.Abbreviation);
            
            var hazardGroups = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.ToList();
            var hazardGroupDictionary = hazardGroups.ToDictionary(key => key.Id, value => value.Name);

            var serverStateIds = serverProfilesAsList.First().Value.Items.Select(item => item.WorkersCompStateId);
            var classCodeByStateDictionary = WorkersCompClassCodesAndHazardsFromBex.GetClassCodeByStateDictionary(serverStateIds);
            
            const int columnCount = 5;
            //at most one
            foreach (var serverProfile in serverProfilesAsList)
            {
                var profile = segment.WorkersCompClassCodeProfile;

                var items = serverProfile.Value.Items;
                var values = new object[items.Count, columnCount];
                var row = 0;
                foreach (var item in items)
                {
                    var stateId = item.WorkersCompStateId;
                    var stateAbbreviation = stateAbbreviations[stateId];
                    var classCodes = classCodeByStateDictionary[stateId];
                    var classCode = classCodes[item.ClassCodeId];
                    var premium = item.Value;

                    values[row, 0] = stateAbbreviation;
                    values[row, 1] = classCode.StateClassCode;
                    values[row, 2] = premium;
                    values[row, 3] = classCode.HazardGroupId.HasValue ? hazardGroupDictionary[classCode.HazardGroupId.Value] : string.Empty;
                    values[row, 4] = classCode.StateDescription;
                    row++;
                }

                var inputRange = profile.ExcelMatrix.GetInputRangeWithoutBuffer();
                if (inputRange.Rows.Count < items.Count)
                {
                    var rowCount = items.Count - inputRange.Rows.Count + 1;
                    var rangeInserter = new WorkersCompClassCodeProfileRowInserter
                    {
                        ExcelRange = inputRange,
                        RowCount = rowCount,
                        StartRow = 3,
                        IsSelectionOnSecondRow = false
                    };
                    using (new ExcelEventDisabler())
                    {
                        rangeInserter.ModifyRange();
                    }
                }

                profile.ExcelMatrix.GetInputRange().Resize[items.Count, columnCount].Value2 = values;
            }
        }

        private static void PopulateWorkersCompRetentionProfile(ISegment segment, IEnumerable<KeyValuePair<long, WorkersCompStateAttachmentDistributionModel>> serverProfiles)
        {
            //at most one
            foreach (var serverProfile in serverProfiles)
            {
                var profile = segment.WorkersCompStateAttachmentProfile;
                var states = StateCodesFromBex.GetWorkersCompStates().OrderBy(state => state.DisplayOrder).ToList();
           
                var stateCounter = 0;
                var stateDictionary = states.ToDictionary(key => Convert.ToInt64(key.Id), value => stateCounter++);

                var distinctAttachments = serverProfile.Value.Items.Select(item => item.Attachment).Distinct().OrderBy(att => att).ToList();
                var attachmentCounter = 0;
                var attachmentDictionary = distinctAttachments.ToDictionary(key => key, value => attachmentCounter++);

                var items = serverProfile.Value.Items;
                var values = new object[states.Count, distinctAttachments.Count];
                foreach (var item in items)
                {
                    var stateId = item.WorkersCompStateId;
                    var attachment = item.Attachment;

                    var row = stateDictionary[stateId];
                    var column = attachmentDictionary[attachment];
                    var value = item.Weight;

                    values[row, column] = value;
                }

                var excelMatrix = profile.ExcelMatrix;
                var inputRange = excelMatrix.GetInputRange();
                excelMatrix.GetInputAttachmentRange().GetFirstColumns(distinctAttachments.Count).Value2 = distinctAttachments.ToArray();
                inputRange.GetFirstColumns(distinctAttachments.Count).Value2 = values;
                excelMatrix.BasisRangeName.GetRange().Value = BexConstants.PercentProfileBasisName;
            }
        }

        private static void MapServerPackageToWorkbook(SubmissionPackageModel serverPackage, IPackage package)
        {
            package.Name = serverPackage.Name;
            package.ResponsibleUnderwriterId = serverPackage.AnalystId;
            package.ResponsibleUnderwriter = UnderwritersFromKeyData.GetName(serverPackage.AnalystId);
            package.SourceId = serverPackage.Id;
            package.SourceTimestamp = DateTime.Now;
            package.CedentId = serverPackage.CedentId;
            
            var keyDataApiWrapperClientFacade = new KeyDataApiWrapperClientFacade(ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.KeyDataBaseUrl);
            var keyDataConverter = new KeyDataConverter(keyDataApiWrapperClientFacade);
            var cedents = keyDataConverter.GetCedents($"0000{serverPackage.CedentId}").ToList();
            package.CedentName = cedents.Any() ? cedents.First().NameAndLocation : string.Empty;

            package.UnderwritingYearExcelMatrix.GetInputRange().Value2 = serverPackage.UnderwritingYear;
            package.IsDirty = false;
        }

        private static Func<IndividualLossModel, DateTimeOffset?> GetCustomOrderBy(ISegment segment)
        {
            var historicPeriodTypeName = HistoricalPeriodTypesFromBex.ReferenceData.First(data => data.Id.ToString() == segment.HistoricalPeriodType).Name;
            Func<IndividualLossModel, DateTimeOffset?> customOrderBy;
            if (historicPeriodTypeName.Contains("Acc"))
            {
                customOrderBy = loss => loss.AccidentDate;
            }
            else if (historicPeriodTypeName.Contains("Pol"))
            {
                customOrderBy = loss => loss.PolicyDate;
            }
            else if (historicPeriodTypeName.Contains("Report"))
            {
                customOrderBy = loss => loss.ReportedDate;
            }
            else
            {
                const string message = "Historic Period Type name not recognized";
                throw new ArgumentOutOfRangeException(message);
            }

            return customOrderBy;
        }
        
        private void MapServerSegmentToSegment(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment)
        {
            var userPreferences = UserPreferences.ReadFromFile();

            var aggregateProperties = GetAggregateProperties(serverPackageId, serverSegment, segment, userPreferences);
            var individualProperties = GetIndividualProperties(serverPackageId, serverSegment, segment, userPreferences);

            segment.SourceId = serverSegment.Id;
            segment.SourceTimestamp = DateTime.Now;
            segment.Name = serverSegment.Name;
            //if(serverSegment.IsUmbrella) segment.IsUmbrella = serverSegment.IsUmbrella;
            segment.HistoricalExposureBasis = serverSegment.HistoricalSubjectBaseUnitOfMeasure.ToString();
            segment.ProspectiveExposureBasis = serverSegment.SubjectBaseUnitOfMeasure.ToString();
            segment.PolicyLimitsApplyTo = serverSegment.SubjectPolicyAlaeTreatment.ToString();
            segment.HistoricalPeriodType = serverSegment.PeriodType.ToString();
            segment.AggregateLossSetDescriptor.IsPaidAvailable = aggregateProperties.HasPaidLoss;
            segment.AggregateLossSetDescriptor.IsLossAndAlaeCombined = aggregateProperties.HasCombinedLoss; 
            segment.IndividualLossSetDescriptor.IsPaidAvailable = individualProperties.HasPaidLoss;
            segment.IndividualLossSetDescriptor.IsLossAndAlaeCombined = individualProperties.HasCombinedLoss;
            segment.IndividualLossSetDescriptor.IsPolicyLimitAvailable = individualProperties.HasPolicyLimit;
            segment.IndividualLossSetDescriptor.IsPolicyAttachmentAvailable = individualProperties.HasPolicyAttachment;
            segment.IndividualLossSetDescriptor.IsEventCodeAvailable = individualProperties.HasEventCode;
            segment.IndividualLossSetDescriptor.IsAccidentDateAvailable = individualProperties.HasAccidentDate;
            segment.IndividualLossSetDescriptor.IsPolicyDateAvailable = individualProperties.HasPolicyDate;
            segment.IndividualLossSetDescriptor.IsReportDateAvailable = individualProperties.HasReportedDate;
            segment.IsDirty = false;
        }

        private IndividualProperties GetIndividualProperties(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment,
            UserPreferences userPreferences)
        {
            IndividualProperties individualProperties;
            if (segment.IndividualLossSets.Any(ils => ils.HasSourceId))
            {
                individualProperties = GetIndividualProperties(serverPackageId, serverSegment);
            }
            else
            {
                segment.IndividualLossSetDescriptor.SetToUserPreferences(userPreferences);
                individualProperties = GetIndividualProperties(segment.IndividualLossSetDescriptor);
            }

            return individualProperties;
        }

        private AggregateProperties GetAggregateProperties(long serverPackageId, SubmissionSegmentModel serverSegment, ISegment segment,
            UserPreferences userPreferences)
        {
            AggregateProperties aggregateProperties = new AggregateProperties();
            if (segment.AggregateLossSets.Any(als => als.HasSourceId))
            {
                aggregateProperties.HasCombinedLoss = GetIsAggregateLossAndAlaeCombined(serverPackageId, serverSegment);
                aggregateProperties.HasPaidLoss = serverSegment.IncludePaidInAggregateLoss;
            }
            else
            {
                segment.AggregateLossSetDescriptor.SetToUserPreferences(userPreferences);
                aggregateProperties.HasCombinedLoss = segment.AggregateLossSetDescriptor.IsLossAndAlaeCombined;
                aggregateProperties.HasPaidLoss = segment.AggregateLossSetDescriptor.IsPaidAvailable;
            }

            return aggregateProperties;
        }

        private bool GetIsAggregateLossAndAlaeCombined(long serverPackageId, SubmissionSegmentModel serverSegment)
        {
            var serverLossSets = _bexCommunicationManager.GetServerAggregateLossSets(serverPackageId, serverSegment.Id);

            foreach (var serverLossSet in serverLossSets)
            {
                var serverLosses = _bexCommunicationManager.GetServerAggregateLosses(serverPackageId, serverSegment.Id, serverLossSet.Key).Values;
                if (serverLosses.Any(loss => loss.ReportedCombinedAmount.HasValue || loss.PaidCombinedAmount.HasValue)) return true;
            }

            return false;
        }

        private IndividualProperties GetIndividualProperties(long serverPackageId, SubmissionSegmentModel serverSegment)
        {
            var individualProperties = new IndividualProperties {HasPaidLoss = serverSegment.IncludePaidInIndividualLoss};

            var serverLossSets = _bexCommunicationManager.GetServerIndividualLossSets(serverPackageId, serverSegment.Id);
            foreach (var serverLossSet in serverLossSets)
            {
                var serverLosses = _bexCommunicationManager.GetServerIndividualLosses(serverPackageId, serverSegment.Id, serverLossSet.Key).Values;
                
                if (serverLosses.Any(loss => loss.ReportedCombinedAmount.HasValue || loss.PaidCombinedAmount.HasValue)) individualProperties.HasCombinedLoss = true;
                if (serverLosses.Any(loss => !string.IsNullOrEmpty(loss.EventCode))) individualProperties.HasEventCode = true;
                if (serverLosses.Any(loss => loss.PolicyLimitAmount.HasValue)) individualProperties.HasPolicyLimit = true;
                if (serverLosses.Any(loss => loss.PolicyAttachmentAmount.HasValue)) individualProperties.HasPolicyAttachment = true;
                if (serverLosses.Any(loss => loss.AccidentDate.HasValue)) individualProperties.HasAccidentDate = true;
                if (serverLosses.Any(loss => loss.PolicyDate.HasValue)) individualProperties.HasPolicyDate = true; 
                if (serverLosses.Any(loss => loss.ReportedDate.HasValue)) individualProperties.HasReportedDate = true;
            }

            return individualProperties;
        }

        private static IndividualProperties GetIndividualProperties(IIndividualLossSetDescriptor individualLossSetDescriptor)
        {
            return new IndividualProperties
            {
                HasPaidLoss = individualLossSetDescriptor.IsPaidAvailable,
                HasCombinedLoss = individualLossSetDescriptor.IsLossAndAlaeCombined,
                HasEventCode = individualLossSetDescriptor.IsEventCodeAvailable,
                HasPolicyLimit = individualLossSetDescriptor.IsPolicyLimitAvailable,
                HasPolicyAttachment = individualLossSetDescriptor.IsPolicyAttachmentAvailable,
                HasAccidentDate = individualLossSetDescriptor.IsAccidentDateAvailable,
                HasPolicyDate = individualLossSetDescriptor.IsPolicyDateAvailable,
                HasReportedDate = individualLossSetDescriptor.IsReportDateAvailable
            };
        }

        private class IndividualProperties
        {
            public bool HasCombinedLoss { get; set; }
            public bool HasEventCode { get; set; }
            public bool HasPolicyLimit { get; set; }
            public bool HasPolicyAttachment { get; set; }
            public bool HasAccidentDate { get; set; }
            public bool HasPolicyDate { get; set; }
            public bool HasReportedDate { get; set; }
            public bool HasPaidLoss { get; set; }
        }

        private class AggregateProperties
        {
            public bool HasCombinedLoss { get; set; }
            public bool HasPaidLoss { get; set; }
        }
    }
}