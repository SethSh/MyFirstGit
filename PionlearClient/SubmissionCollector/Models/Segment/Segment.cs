using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using PionlearClient.Model;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelWorkspaceFolder;
using SubmissionCollector.Models.Comparers;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Enums;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.Models.Subline;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel.ItemSources;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace SubmissionCollector.Models.Segment
{
    [JsonObject]
    public class Segment : BaseInventoryItem, ISegment, ISegmentInventoryItem
    {
        private string _name;
        private string _prospectiveExposureBasis;
        private string _historicalExposureBasis;
        private string _policyLimitsApplyTo;
        private string _historicalPeriodType;
        private bool _isUmbrella;
    
        public Segment()
        {
            WorksheetManager = new WorksheetManager(this);
        }

        public Segment(bool isNotCalledFromJson)
        {
            try
            {
                IsCurrentlyCreating = true;

                var userPreferences = UserPreferences.ReadFromFile();

                Id = GetNextId();
                Name = $"Segment {Id + 1}";
                Guid = Guid.NewGuid();
                DisplayOrder = GetNextDisplayOrder();
                ProspectiveExposureBasis = userPreferences.ExposureBasis.ToString();
                HistoricalExposureBasis = userPreferences.ExposureBasis.ToString();
                HistoricalPeriodType = userPreferences.PeriodType.ToString();
                PolicyLimitsApplyTo = userPreferences.SubjectPolicyAlaeTreatment.ToString();
                
                ExcelComponents = new List<IExcelComponent> {new SublineProfile(Id), new UmbrellaProfile(Id), new PeriodSet(Id)};

                AggregateLossSetDescriptor = new AggregateLossSetDescriptor(this);
                AggregateLossSetDescriptor.SetToUserPreferences(userPreferences);
                IndividualLossSetDescriptor = new IndividualLossSetDescriptor(this);
                IndividualLossSetDescriptor.SetToUserPreferences(userPreferences);
                IsExpanded = true;

                OrphanExcelMatrices = new List<ISegmentExcelMatrix> {new ProspectiveExposureAmountExcelMatrix(Id)};

                SublineExcelMatrix.ProfileFormatter = ProfileFormatterFactory.Create(userPreferences.ProfileBasisId);
                UmbrellaExcelMatrix.ProfileFormatter = ProfileFormatterFactory.Create(userPreferences.ProfileBasisId);
                _sublines = new List<ISubline>();
                
                IsWorkerCompClassCodeActive = userPreferences.WorkersCompClassCodesIsActive;

                WorksheetManager = new WorksheetManager(this);
                RegisterBaseChangeEvents();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                IsCurrentlyCreating = false;
            }
        }

        [JsonIgnore]
        [Browsable(false)]
        public static bool IsCurrentlyDuplicating { get; set; }

        [JsonIgnore]
        [Browsable(false)]
        public static bool IsCurrentlyFixing { get; set; }

        [JsonIgnore]
        [Browsable(false)]
        public static bool IsCurrentlyCreating { get; set; }

        [JsonProperty]
        private readonly IList<ISubline> _sublines;

        [PropertyOrder(0)]
        [Category("Attributes")]
        [DisplayName(@"Submission Segment Name")]
        [Description("Name of submission segment and worksheet")]
        [Browsable(true)]
        [ReadOnly(false)]
        public new string Name
        {
            get => _name;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || IsCurrentlyDuplicating)
                {
                    _name = value;
                    NotifyPropertyChanged();
                    return;
                }

                if (!ValidateName(value)) return;

                _name = value;
                NotifyPropertyChanged();
                ChangeWorksheet();
            }
        }

        [Browsable(false)]
        public int Id { get; set; }

        [PropertyOrder(200)]
        [ItemsSource(typeof(ExposureBasisSource))]
        [Category("Attributes")]
        [DisplayName(@"Historical Exposure Basis")]
        [Description("Historical exposure basis can't be changed once this submission segment is 'used by' a BEX rating analysis.")]
        [Browsable(true)]
        [ReadOnly(false)]
        public string HistoricalExposureBasis
        {
            get => _historicalExposureBasis;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || IsCurrentlyDuplicating)
                {
                    _historicalExposureBasis = value;
                    NotifyPropertyChanged();
                    return;
                }

                using (new ExcelEventDisabler())
                {
                    WorksheetManager?.SetHistoricalExposureSetsBasis(ExposureBasisFromBex.GetExposureBasisName(Convert.ToInt32(value)));
                }
                _historicalExposureBasis = value;
                NotifyPropertyChanged();
            }
        }

        [PropertyOrder(300)]
        [JsonProperty("PropspectiveExposureBasis")]
        [ItemsSource(typeof(ExposureBasisSource))]
        [Category("Attributes")]
        [DisplayName(@"Prospective Exposure Basis")]
        [Description("Prospective exposure basis can't be changed once this submission segment is 'used by' a BEX rating analysis.")]
        [Browsable(true)]
        [ReadOnly(false)]
        public string ProspectiveExposureBasis
        {
            get => _prospectiveExposureBasis;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || IsCurrentlyDuplicating)
                {
                    _prospectiveExposureBasis = value;
                    NotifyPropertyChanged();
                    return;
                }

                WorksheetManager?.SetProspectiveExposureBasis(ExposureBasisFromBex.GetExposureBasisName(Convert.ToInt16(value)));
                _prospectiveExposureBasis = value;
                NotifyPropertyChanged();

                if (WorksheetManager == null) return;

                var summaryBuilder = new ProspectiveExposureSummaryBuilder();
                summaryBuilder.Build();
            }
        }
        
        [JsonIgnore]
        [Browsable(false)]
        public IEnumerable<IInventoryItem> SublineViews => _sublines.OfType<IInventoryItem>().ToList();

        [PropertyOrder(100)]
        [Browsable(true)]
        [ReadOnly(false)]
        [Category("Attributes")]
        [DisplayName(@"Umbrella")]
        [Description("Is the submission segment umbrella.  The umbrella decision can't be changed once this submission segment is 'used by' a BEX analysis.")]
        public bool IsUmbrella
        {
            get => _isUmbrella;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || IsCurrentlyDuplicating || IsCurrentlyRebuilding)
                {
                    _isUmbrella = value;
                    NotifyPropertyChanged();
                    return;
                }

                if (!IsStructureModifiable)
                {
                    MessageHelper.Show(GetBlockModificationsMessage(BexConstants.UmbrellaTypeName), MessageType.Stop);
                    return;
                }
                
                if (!value)
                {
                    if (!AreUmbrellaTypesModifiable)
                    {
                        var message = $"In order to turn off {BexConstants.UmbrellaName.ToLower()}, " +
                                      $"all commercial {BexConstants.PolicyProfileName.ToLower()}s must treat {BexConstants.UmbrellaTypeName.ToLower()} as grouped.";
                        MessageHelper.Show(message, MessageType.Stop);
                        return;
                    }
                }

                if (value && IsWorkersComp)
                {
                    var message = $"{BexConstants.UmbrellaName.ToStartOfSentence()} isn't valid for {BexConstants.WorkersCompName}";
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }

                if (value && IsProperty)
                {
                    var message = $"{BexConstants.UmbrellaName.ToStartOfSentence()} isn't valid for {BexConstants.PropertyName}";
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }

                if (value && ContainsAutoPhysicalDamage())
                {
                    var message = $"{BexConstants.UmbrellaName.ToStartOfSentence()} {BexConstants.SegmentName.ToLower()} " +
                                  $"can't include an auto physical damage {BexConstants.SublineName.ToLower()}";
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }
                
                using (new ExcelEventDisabler())
                {
                    using (new ExcelScreenUpdateDisabler())
                    {
                        if (value && ContainsAnyCommercialSublines && !IsCurrentlyRebuilding)
                        {
                            var umbrellaSelectorManager = new UmbrellaTypeAllocatorManager();
                            var response = umbrellaSelectorManager.SelectUmbrellaTypes();
                            if (response == FormResponse.Cancel) return;
                        }
                        else
                        {
                            if (UmbrellaExcelMatrix.RangeName.ExistsInWorkbook())
                            {
                                WorksheetManager.DeleteUmbrellaMatrix();
                            }
                        }
                    }
                }

                _isUmbrella = value;
                NotifyPropertyChanged();
                ExcelSheetActivateEventManager.RefreshUmbrellaWizardButton(this);
            }
        }

        [JsonIgnore]
        [Browsable(false)]
        public bool IsCurrentlyRebuilding { get; set; }

        [Browsable(false)]
        public bool IsWorkerCompClassCodeActive { get; set; }

        [PropertyOrder(800)]
        [DisplayName(@"Subject Policy Limits Apply To")]
        [Browsable(true)]
        [ReadOnly(false)]
        [Category("Attributes")]
        [ItemsSource(typeof(PolicyLimitsApplyToSource))]
        public string PolicyLimitsApplyTo
        {
            get => _policyLimitsApplyTo;
            set
            {
                _policyLimitsApplyTo = value;
                NotifyPropertyChanged();
            }
        }

        [Browsable(false)]
        [JsonIgnore]
        public IPackage ParentPackage => Globals.ThisWorkbook.ThisExcelWorkspace.Package;

        [ItemsSource(typeof(HistoricalPeriodTypeSource))]
        [DisplayName(@"Historical Period")]
        [Browsable(true)]
        [ReadOnly(false)]
        [Category("Attributes")]
        [PropertyOrder(900)]
        public string HistoricalPeriodType
        {
            get => _historicalPeriodType;
            set
            {
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook || IsCurrentlyDuplicating)
                {
                    _historicalPeriodType = value;
                    NotifyPropertyChanged();
                    return;
                }
                
                WorksheetManager?.SetHistoricalPeriodType(HistoricalPeriodTypesFromBex.GetName(Convert.ToInt32(value)));
                _historicalPeriodType = value;
                NotifyPropertyChanged();
            }
        }
        
        [DisplayName(@"Aggregate Loss")]
        [Category("Attributes")]
        [PropertyOrder(1000)]
        [ExpandableObject]
        public IAggregateLossSetDescriptor AggregateLossSetDescriptor { get; set; }

        [DisplayName(@"Individual Loss")]
        [Category("Attributes")]
        [PropertyOrder(1001)]
        [ExpandableObject]
        public IIndividualLossSetDescriptor IndividualLossSetDescriptor { get; set; }

        [Browsable(false)]
        public WorksheetManager WorksheetManager { get; set; }
        
        [Browsable(false)]
        public string HeaderRangeName => $"segment{Id}.{ExcelConstants.HeaderRangeName}";

        [Browsable(false)]
        public bool ContainsAnyCommercialSublines => this.Any(subline => !subline.IsPersonal);

        [Browsable(false)]
        public bool ContainsAnyPersonalSublines => this.Any(subline => subline.IsPersonal);

        [Browsable(false)]
        public bool IsLiability => !IsProperty && !IsWorkersComp;
        
        [Browsable(false)]
        public bool IsProperty => this.All(x => x.LineOfBusinessType == LineOfBusinessType.Property);

        [Browsable(false)]
        public bool IsWorkersComp => this.All(subline => subline.LineOfBusinessType == LineOfBusinessType.WorkersCompensation);

        [Browsable(false)]
        public List<IExcelComponent> ExcelComponents { get; set; }
        
        [Browsable(false)]
        [JsonIgnore]
        public SublineProfile SublineProfile => ExcelComponents.OfType<SublineProfile>().Single();

        [Browsable(false)]
        [JsonIgnore]
        public UmbrellaProfile UmbrellaProfile => ExcelComponents.OfType<UmbrellaProfile>().Single();

        
        [Browsable(false)]
        [JsonIgnore]
        public WorkersCompStateHazardGroupProfile WorkersCompStateHazardGroupProfile => ExcelComponents.OfType<WorkersCompStateHazardGroupProfile>().SingleOrDefault();

        [Browsable(false)]
        [JsonIgnore]
        public WorkersCompClassCodeProfile WorkersCompClassCodeProfile => ExcelComponents.OfType<WorkersCompClassCodeProfile>().SingleOrDefault();

        [Browsable(false)]
        [JsonIgnore]
        public WorkersCompStateAttachmentProfile WorkersCompStateAttachmentProfile => ExcelComponents.OfType<WorkersCompStateAttachmentProfile>().SingleOrDefault();

        [Browsable(false)]
        [JsonIgnore]
        public MinnesotaRetention WorkersCompMinnesotaRetention => ExcelComponents.OfType<MinnesotaRetention>().SingleOrDefault();


        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<IProfile> Profiles => ExcelComponents.OfType<IProfile>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<IHistorical> Historicals => ExcelComponents.OfType<IHistorical>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<PolicyProfile> PolicyProfiles => ExcelComponents.OfType<PolicyProfile>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<HazardProfile> HazardProfiles => ExcelComponents.OfType<HazardProfile>();
        
        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<StateProfile> StateProfiles => ExcelComponents.OfType<StateProfile>();

        
        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<ConstructionTypeProfile> ConstructionTypeProfiles => ExcelComponents.OfType<ConstructionTypeProfile>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<OccupancyTypeProfile> OccupancyTypeProfiles => ExcelComponents.OfType<OccupancyTypeProfile>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<ProtectionClassProfile> ProtectionClassProfiles => ExcelComponents.OfType<ProtectionClassProfile>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<TotalInsuredValueProfile> TotalInsuredValueProfiles => ExcelComponents.OfType<TotalInsuredValueProfile>();

        
        [Browsable(false)]
        [JsonIgnore]
        public PeriodSet PeriodSet => ExcelComponents.OfType<PeriodSet>().Single();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<ExposureSet> ExposureSets => ExcelComponents.OfType<ExposureSet>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<AggregateLossSet> AggregateLossSets => ExcelComponents.OfType<AggregateLossSet>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<IndividualLossSet> IndividualLossSets => ExcelComponents.OfType<IndividualLossSet>();

        [Browsable(false)]
        [JsonIgnore]
        public IEnumerable<RateChangeSet> RateChangeSets => ExcelComponents.OfType<RateChangeSet>();

        [Browsable(false)]
        [JsonIgnore]
        public SegmentModel SegmentModel { get; set; }

        [Browsable(false)]
        [JsonIgnore]
        public IList<ISegmentExcelMatrix> ExcelMatrices
        {
            get
            {
                var list = new List<ISegmentExcelMatrix>();

                list.AddRange(OrphanExcelMatrices);
                list.AddRange(ExcelComponents.Select(x => x.CommonExcelMatrix));

                return list;
            }
        }

        [Browsable(false)]
        public IList<ISegmentExcelMatrix> OrphanExcelMatrices { get; set; }

        [Browsable(false)]
        [JsonIgnore]
        public ProspectiveExposureAmountExcelMatrix ProspectiveExposureAmountExcelMatrix => ExcelMatrices.OfType<ProspectiveExposureAmountExcelMatrix>().Single();

        [Browsable(false)]
        [JsonIgnore]
        public string ProspectiveExposureAmountExcelCell => ProspectiveExposureAmountExcelMatrix.GetInputRange().Address[External: true];

        [Browsable(false)]
        [JsonIgnore]
        public SublineExcelMatrix SublineExcelMatrix => SublineProfile.SublineExcelMatrix;

        [Browsable(false)]
        [JsonIgnore]
        public UmbrellaExcelMatrix UmbrellaExcelMatrix => UmbrellaProfile.UmbrellaExcelMatrix;
        
        [Browsable(false)]
        public int Count => _sublines.Count;

        [Browsable(false)]
        public bool IsReadOnly => false;

        [Browsable(false)]
        [JsonIgnore]
        public bool IsStructureModifiable => !SourceId.HasValue || ParentPackage.AttachedToRatingAnalysisStatus == AttachOptions.False;

        [Browsable(false)]
        [JsonIgnore]
        public bool AreUmbrellaTypesModifiable
        {
            get
            {
                var personalCode = UmbrellaTypesFromBex.GetPersonalCode();
                return !PolicyProfiles.Any(policyProfile =>
                    policyProfile.UmbrellaType.HasValue && policyProfile.UmbrellaType.Value != personalCode);
            }
        }


        public string GetBlockModificationsMessage(string type)
        {
            var sb = new StringBuilder();
            sb.AppendLine(ParentPackage.AttachedToRatingAnalysisMessage);
            sb.AppendLine();
            sb.AppendLine($"{type.ToStartOfSentence()} {BexConstants.UpdateName.ToLower()}s to " +
                          $"{BexConstants.SegmentName.ToLower()} <{Name}> is blocked.");
            return sb.ToString();
        }

        public ISegment Duplicate()
        {
            try
            {
                IsCurrentlyDuplicating = true;
                var duplicate = this.DeepClone();

                duplicate.Id = GetNextId();
                duplicate.Name = Globals.ThisWorkbook.MakeWorksheetNameAcceptable(Name);
                duplicate.Guid = Guid.NewGuid();
                duplicate.IsDirty = true;
                duplicate.IsSelected = true;
                duplicate.SourceId = null;
                duplicate.SourceTimestamp = null;
                foreach (var item in duplicate)
                {
                    item.SegmentId = duplicate.Id;
                }
                
                foreach (var excelComponent in duplicate.ExcelComponents)
                {
                    excelComponent.SourceId = null;
                    excelComponent.SourceTimestamp = null;
                    excelComponent.IsDirty = true;
                    excelComponent.Guid = Guid.NewGuid();
                }

                foreach (var duplicateExcelComponent in duplicate.ExcelComponents.OfType<IProvidesLedger>())
                {
                    duplicateExcelComponent.Ledger.Clear();
                }

                duplicate.WorksheetManager.DuplicateWorksheet(this);

                foreach (var ec in duplicate.ExcelComponents)
                {
                    ec.SegmentId = duplicate.Id;
                    ec.CommonExcelMatrix.SegmentId = duplicate.Id;
                }
                duplicate.OrphanExcelMatrices.ForEach(oem => oem.SegmentId = duplicate.Id);
                
                IsCurrentlyDuplicating = false;
                return duplicate;
            }
            finally
            {
                IsCurrentlyDuplicating = false;
            }
        }

        public override void DecoupleFromServer()
        {
            base.DecoupleFromServer();
            ExcelComponents.ForEach(ec =>
            {
                ec.DecoupleFromServer();
                DecoupleManager.ClearLedger(ec);
            });
        }

        public void UpdateSublines(IList<ISubline> sublines)
        {
            var list = this.Except(sublines).ToList();
            list.ForEach(x => Remove(x));
            foreach (var subline in sublines.Except(this))
            {
                Add(subline);
            }
        }

        public StringBuilder CreateSegmentModelForQualityControl()
        {
            var qualityControlMessages = new StringBuilder();

            var segmentValidations = Validate();
            if (segmentValidations.Length > 0)
            {
                qualityControlMessages.Append(segmentValidations);
            }

            CreatePolicyProfileModels(qualityControlMessages);
            CreateStateProfileModels(qualityControlMessages);
            CreateHazardProfileModels(qualityControlMessages);

            CreateConstructionTypeProfileModels(qualityControlMessages);
            CreateProtectionClassProfileModels(qualityControlMessages);
            CreateOccupancyTypeProfileModels(qualityControlMessages);
            CreateTotalInsuredValueProfileModels(qualityControlMessages);

            CreateWorkersCompStateAttachmentProfileModel(qualityControlMessages);
            CreateWorkersCompClassCodeProfileModel(qualityControlMessages);
            CreateWorkersCompStateHazardProfileModel(qualityControlMessages);
            CreateWorkersCompMinnesotaRetention(qualityControlMessages);

            CreateValidPeriods(qualityControlMessages);
            CreateAggregateLossModels(qualityControlMessages);
            CreateIndividualLossModels(qualityControlMessages);
            CreateRateChangeModels(qualityControlMessages);

            return qualityControlMessages;
        }

        
        public StringBuilder CreateSegmentModel()
        {
            var validations = new StringBuilder();

            var segmentValidations = Validate();
            if (segmentValidations.Length > 0)
            {
                validations.Append(segmentValidations);
            }
            
            CreatePolicyProfileModels(validations);
            CreateStateProfileModels(validations);
            CreateHazardProfileModels(validations);

            CreateConstructionTypeProfileModels(validations);
            CreateProtectionClassProfileModels(validations);
            CreateOccupancyTypeProfileModels(validations);
            CreateTotalInsuredValueProfileModels(validations);

            CreateWorkersCompStateAttachmentProfileModel(validations);
            CreateWorkersCompClassCodeProfileModel(validations);
            CreateWorkersCompStateHazardProfileModel(validations);
            CreateWorkersCompMinnesotaRetention(validations);

            CreateValidPeriods(validations);
            CreateExposureSetModels(validations);
            CreateAggregateLossModels(validations);
            CreateIndividualLossModels(validations);
            CreateRateChangeModels(validations);

            //SegmentModel.IndividualLossSetModels isn't created until CreateIndividualLossModels
            ValidateSubjectPolicyAppliesTo(validations);

            return validations;
        }

        public void SetAllLedgerItemsToNotDirty()
        {
            ExposureSets.ForEach(set => set.Ledger.SetToNotDirty());
            AggregateLossSets.ForEach(set => set.Ledger.SetToNotDirty());
            IndividualLossSets.ForEach(set => set.Ledger.SetToNotDirty());
            RateChangeSets.ForEach(set => set.Ledger.SetToNotDirty());
        }

        public bool IsNameAcceptableLength()
        {
            if (Name.Length <= ExcelConstants.WorksheetNameMaximumCharacters - 3) return true;

            var message = $"{BexConstants.SegmentName} duplicating requires a shorter submission segment name.  Trim your {BexConstants.SegmentName.ToLower()} name and try again.";
            MessageHelper.Show(message, MessageType.Stop);
            return false;
        }

        private void ValidateSubjectPolicyAppliesTo(StringBuilder validations)
        {
            var anyIndividualLoss = SegmentModel.IndividualLossSetModels.Any(model => model.Items.Any());
            if (!anyIndividualLoss) return;

            var policyLimitApplyToCode = Convert.ToInt16(PolicyLimitsApplyTo);
            var policyLimitAppliesToLossOnly = policyLimitApplyToCode == BexConstants.LossOnlySubjectPolicyAlaeTreatmentId;
            var hasPolicyLimits = IndividualLossSetDescriptor.IsPolicyLimitAvailable
                                  && IndividualLossSets.Any(ils => ils.ExcelMatrix.Items.Any(item => item.PolicyLimitAmount.HasValue));

            
            if (IndividualLossSetDescriptor.IsLossAndAlaeCombined 
                && policyLimitAppliesToLossOnly 
                && hasPolicyLimits)
            {
                validations.Append($"Can't have combined {BexConstants.IndividualLossName.ToLower()} " +
                                   $"and subject {BexConstants.PolicyLimitName.ToLower()}s apply to loss only " +
                                   $"and {BexConstants.IndividualLossName.ToLower()} {BexConstants.PolicyLimitName.ToLower()}s.");
            }
        }

        private void CreateValidPeriods(StringBuilder qualityControlMessages)
        {
            var periodsValidation = PeriodSet.ExcelMatrix.Validate();
            if (periodsValidation.Length > 0)
            {
                qualityControlMessages.AppendLine(BexConstants.PeriodSetName);
                qualityControlMessages.Append(periodsValidation);
                qualityControlMessages.AppendLine();
            }
            SegmentModel.ValidPeriods = PeriodSet.ExcelMatrix.Items;
        }

        private void CreateHazardProfileModels(StringBuilder validations)
        {
            foreach (var profile in HazardProfiles)
            {
                var model = (HazardModel)profile.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreateStateProfileModels(StringBuilder validations)
        {
            foreach (var profile in StateProfiles)
            {
                var model = (StateModel)profile.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreatePolicyProfileModels(StringBuilder validations)
        {
            foreach (var profile in PolicyProfiles)
            {
                var model = (PolicyModel)profile.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreateConstructionTypeProfileModels(StringBuilder validations)
        {
            foreach (var profile in ConstructionTypeProfiles)
            {
                var model = (ConstructionTypeModel)profile.CreateModel(validations);
                SegmentModel.Models.Add(model); 
            }
        }

        private void CreateOccupancyTypeProfileModels(StringBuilder validations)
        {
            foreach (var profile in OccupancyTypeProfiles)
            {
                var model = (OccupancyTypeModel)profile.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreateProtectionClassProfileModels(StringBuilder validations)
        {
            foreach (var profile in ProtectionClassProfiles)
            {
                var model = (ProtectionClassModel)profile.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreateTotalInsuredValueProfileModels(StringBuilder validations)
        {
            foreach (var profile in TotalInsuredValueProfiles)
            {
                var model = (TotalInsuredValueModel)profile.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }


        private void CreateWorkersCompStateAttachmentProfileModel(StringBuilder validations)
        {
            if (WorkersCompStateAttachmentProfile == null) return;
            
            var model = (WorkersCompStateAttachmentModel)WorkersCompStateAttachmentProfile.CreateModel(validations);
            SegmentModel.Models.Add(model);
        }

        private void CreateWorkersCompClassCodeProfileModel(StringBuilder validations)
        {
            if (WorkersCompClassCodeProfile == null) return;

            var model = (WorkersCompClassCodeModel)WorkersCompClassCodeProfile.CreateModel(validations);
            SegmentModel.Models.Add(model);
        }


        private void CreateWorkersCompStateHazardProfileModel(StringBuilder validations)
        {
            if (WorkersCompStateHazardGroupProfile == null) return;

            var model = (WorkersCompStateHazardGroupModel)WorkersCompStateHazardGroupProfile.CreateModel(validations);
            SegmentModel.Models.Add(model);
        }


        private void CreateWorkersCompMinnesotaRetention(StringBuilder validations)
        {
            if (WorkersCompMinnesotaRetention == null) return;

            var model = (MinnesotaRetentionModel)WorkersCompMinnesotaRetention.CreateModel(validations); 
            SegmentModel.MinnesotaRetentionId = model.RetentionId;
        }

        private void CreateExposureSetModels(StringBuilder validations)
        {
            foreach (var exposureSet in ExposureSets)
            {
                var model = (ExposureSetModel)exposureSet.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreateAggregateLossModels(StringBuilder validations)
        {
            foreach (var lossSet in AggregateLossSets)
            {
                var model = (AggregateLossSetModel)lossSet.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreateIndividualLossModels(StringBuilder validations)
        {
            foreach (var lossSet in IndividualLossSets)
            {
                var model = (IndividualLossSetModel)lossSet.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }

        private void CreateRateChangeModels(StringBuilder validations)
        {
            foreach (var rateChangeSet in RateChangeSets)
            {
                var model = (RateChangeSetModel)rateChangeSet.CreateModel(validations);
                SegmentModel.Models.Add(model);
            }
        }
        
        private StringBuilder Validate()
        {
            var validations = new StringBuilder();
            SegmentModel = MapToModel();

            var prospectiveExposureAmountValidations = ProspectiveExposureAmountExcelMatrix.Validate();
            if (prospectiveExposureAmountValidations.Length > 0)
            {
                validations.Append(prospectiveExposureAmountValidations);
                validations.AppendLine();
            }
            SegmentModel.ProspectiveExposureAmount = ProspectiveExposureAmountExcelMatrix.Item;

            var sublineAllocationsValidations = SublineExcelMatrix.Validate();
            if (sublineAllocationsValidations.Length > 0)
            {
                validations.Append(sublineAllocationsValidations);
                validations.AppendLine();
            }
            
            SegmentModel.ValidSublines = SublineExcelMatrix.Allocations;
            //Phase 2 allow units
            //SegmentModel.IsSublineAllocationPercent = SublineProfile.SublineExcelMatrix.ProfileFormatter is PercentProfileFormatter;
            SegmentModel.IsSublineAllocationPercent = true;

            if (IsUmbrella)
            {
                var umbrellaValidations = UmbrellaExcelMatrix.Validate();
                if (umbrellaValidations.Length > 0)
                {
                    validations.Append(umbrellaValidations);
                    validations.AppendLine();
                }
                SegmentModel.UmbrellaTypeAllocations = UmbrellaExcelMatrix.Allocations;
            }

            #region first - check for empty sublines
            var excelMatrices = ExcelMatrices.OfType<MultipleOccurrenceSegmentExcelMatrix>();
            excelMatrices.ForEach(matrix =>
            {
                if (matrix.Count != 0) return;
                validations.AppendLine($"{matrix.FriendlyName.ToStartOfSentence()} contains no sublines");
            });
            #endregion

            #region Second - check for unassigned sublines
            var policyProfilesSublines = new List<ISubline>();
            PolicyProfiles.ForEach(p => policyProfilesSublines.AddRange(p.ExcelMatrix));
            var unassignedSublines = this.Where(x => x.HasPolicyProfile).Except(policyProfilesSublines, new SublineComparer());
            unassignedSublines.ForEach(x => validations.AppendLine($"No {BexConstants.PolicyProfileName.ToLower()} contains {x.ShortNameWithLob}"));

            var stateProfilesSublines = new List<ISubline>();
            StateProfiles.ForEach(s => stateProfilesSublines.AddRange(s.ExcelMatrix));
            unassignedSublines = this.Where(x => x.HasStateProfile).Except(stateProfilesSublines, new SublineComparer());
            unassignedSublines.ForEach(x => validations.AppendLine($"No {BexConstants.StateProfileName.ToLower()} contains {x.ShortNameWithLob}"));

            var historicalExposureSetsSublines = new List<ISubline>();
            ExposureSets.ForEach(e => historicalExposureSetsSublines.AddRange(e.ExcelMatrix));
            unassignedSublines = this.Except(historicalExposureSetsSublines, new SublineComparer());
            unassignedSublines.ForEach(x => validations.AppendLine($"No {BexConstants.ExposureSetName.ToLower()} contains {x.ShortNameWithLob}"));

            var aggregateReportedLossSetsSublines = new List<ISubline>();
            AggregateLossSets.ForEach(l => aggregateReportedLossSetsSublines.AddRange(l.ExcelMatrix));
            unassignedSublines = this.Except(aggregateReportedLossSetsSublines, new SublineComparer());
            unassignedSublines.ForEach(x => validations.AppendLine($"No {BexConstants.AggregateLossSetName.ToLower()} contains {x.ShortNameWithLob}"));

            var individualLossSetsSublines = new List<ISubline>();
            IndividualLossSets.ForEach(l => individualLossSetsSublines.AddRange(l.ExcelMatrix));
            unassignedSublines = this.Except(individualLossSetsSublines, new SublineComparer());
            unassignedSublines.ForEach(x => validations.AppendLine($"No {BexConstants.IndividualLossSetName.ToLower()} contains {x.ShortNameWithLob}"));
            #endregion

            ValidateMissingIndividualDateFields(validations);

            return validations;
        }

        private SegmentModel MapToModel()
        {
            return SegmentModel = new SegmentModel
            {
                SourceId = SourceId,
                Guid = Guid,
                ProspectiveExposureBasisCode = Convert.ToInt16(ProspectiveExposureBasis),
                HistoricalExposureBasisCode = Convert.ToInt16(HistoricalExposureBasis),
                HistoricalPeriodType = Convert.ToInt16(HistoricalPeriodType),
                IsUmbrella = IsUmbrella,
                Name = Name,
                SubjectPolicyAlaeTreatment = Convert.ToInt16(PolicyLimitsApplyTo),
                IncludePaidInAggregateLoss = AggregateLossSetDescriptor.IsPaidAvailable,
                IncludePaidInIndividualLoss = IndividualLossSetDescriptor.IsPaidAvailable,
                IsProperty = IsProperty,
                IsDirty = IsDirty
            };
        }

        private void ValidateMissingIndividualDateFields(StringBuilder validations)
        {
            string missingDateName = null;
            switch (SegmentModel.HistoricalPeriodType)
            {
                case 1:
                    if (!IndividualLossSetDescriptor.IsAccidentDateAvailable)
                    {
                        missingDateName = $"{BexConstants.AccidentDateName}";
                    }
                    break;

                case 2:
                    if (!IndividualLossSetDescriptor.IsPolicyDateAvailable)
                    {
                        missingDateName = $"{BexConstants.PolicyDateName}";
                    }
                    break;

                case 3:
                    if (!IndividualLossSetDescriptor.IsReportDateAvailable)
                    {
                        missingDateName = $"{BexConstants.ReportDateName}";
                    }
                    break;

                default: throw new ArgumentOutOfRangeException($"Can't find {BexConstants.PeriodName.ToLower()} type");
            }

            if (string.IsNullOrEmpty(missingDateName)) return;
            
            var dateName = HistoricalPeriodTypesFromBex.GetName(SegmentModel.HistoricalPeriodType);
            validations.AppendLine($"Change either {BexConstants.PeriodName.ToLower()} <{dateName}> or select {BexConstants.IndividualLossSetName.ToLower()} <{missingDateName}>");
            validations.AppendLine();
        }

        private static int GetNextId()
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            return package.Segments.Any() ? package.Segments.Select(p => p.Id).Max() + 1 : 0;
        }

        private static int GetNextDisplayOrder()
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            return package.Segments.Any() ? package.Segments.Select(p => p.DisplayOrder).Max() + 1 : 0;
        }

        private void ChangeWorksheet()
        {
            if (WorksheetManager == null) return;
            using (new ExcelEventDisabler())
            {
                Globals.ThisWorkbook.WriteSegmentNameToHeader(this);
            }
            ChangeSegmentWorksheetName();
        }

        private bool ValidateName(string candidate)
        {
            const string title = BexConstants.DataValidationTitle;
            var lowerSegmentName = BexConstants.SegmentName.ToLower();

            if (string.IsNullOrEmpty(candidate))
            {
                MessageHelper.Show(title, "Can't be blank", MessageType.Stop);
                return false;
            }

            if (Name == candidate) return false;

            if (Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Any(x => x.Name.Equals(candidate,
                StringComparison.OrdinalIgnoreCase)))
            {
                MessageHelper.Show(title, "Name already exists and duplicates aren't allowed", MessageType.Stop);
                return false;
            }

            const string because = "Because submission segment name will be used for worksheet name";
            if (Globals.ThisWorkbook.DoesWorksheetNameExceedMaximumLength(candidate))
            {
                MessageHelper.Show(title,
                    $"{because}, {lowerSegmentName} " +
                    $"name length can't exceed {ExcelConstants.WorksheetNameMaximumCharacters} characters",
                    MessageType.Stop);
                return false;
            }

            if (Globals.ThisWorkbook.DoesWorksheetNameExist(candidate))
            {
                MessageHelper.Show(title, $"{because}, {lowerSegmentName} name can't duplicate an existing worksheet name", MessageType.Stop);
                return false;
            }

            if (Globals.ThisWorkbook.DoesWorksheetNameContainInvalidCharacters(candidate))
            {
                MessageHelper.Show(title, $"{because}, {lowerSegmentName} name can't contain invalid characters", MessageType.Stop);
                return false;
            }
            return true;
        }

        private void ChangeSegmentWorksheetName()
        {
            if (IsWorksheetNameNotOkay()) return;
            using (new WorkbookUnprotector())
            {
                WorksheetManager.Worksheet.Name = Name;
            }
        }

        private bool IsWorksheetNameNotOkay()
        {
            return Globals.ThisWorkbook.DoesWorksheetNameExceedMaximumLength(Name)
                   || Globals.ThisWorkbook.DoesWorksheetNameExist(Name)
                   || Globals.ThisWorkbook.DoesWorksheetNameContainInvalidCharacters(Name);
        }
        
        private bool ContainsAutoPhysicalDamage()
        {
            return this.Any(x => x.LineOfBusinessType == LineOfBusinessType.PhysicalDamage);
        }


        public void AddWorkersCompStateHazardGroupProfile()
        {
            if (!IsWorkersComp || WorkersCompStateHazardGroupProfile != null) return;

            var profile = new WorkersCompStateHazardGroupProfile(Id)
            {
                Name = $"{BexConstants.WorkersCompStateHazardGroupName}",
            };

            profile.ExcelMatrix.IsIndependent = !IsWorkerCompClassCodeActive && UserPreferences.ReadFromFile().WorkersCompStateByHazardGroupIsIndependent;
            ExcelComponents.Add(profile);
        }

        public void AddWorkersCompClassCodeProfile()
        {
            if (!IsWorkersComp || WorkersCompClassCodeProfile != null) return;
            if (!IsWorkerCompClassCodeActive) return;

            var workersCompClassCodeProfile = new WorkersCompClassCodeProfile(Id)
            {
                Name = $"{BexConstants.WorkersCompClassCodeName}",
            };

            ExcelComponents.Add(workersCompClassCodeProfile);
        }

        public void AddWorkersCompStateAttachmentProfile()
        {
            if (!IsWorkersComp || WorkersCompStateAttachmentProfile != null) return;

            var workersCompStateAttachmentProfile = new WorkersCompStateAttachmentProfile(Id)
            {
                Name = $"{BexConstants.WorkersCompStateAttachmentName}",
            };

            ExcelComponents.Add(workersCompStateAttachmentProfile);
        }


        public void AddWorkersCompMinnesotaRetention()
        {
            if (!IsWorkersComp || WorkersCompMinnesotaRetention != null) return;

            var minnesotaRetention = new MinnesotaRetention(Id)
            {
                Name = $"{BexConstants.MinnesotaRetentionName}",
            };

            ExcelComponents.Add(minnesotaRetention);
        }



        public void AddProtectionClassProfile(ISubline item)
        {
            if (!item.IsProperty()) return;

            var displayOrder = ProtectionClassProfiles.Any()
                ? ProtectionClassProfiles.Max(x => x.IntraDisplayOrder) + 1
                : 0;

            var protectionClassProfile = new ProtectionClassProfile(Id)
            {
                Name = $"{BexConstants.ProtectionClassName} {item.ShortName}",
                ComponentId = item.Code
            };
            protectionClassProfile.ExcelMatrix.Add(item);
            protectionClassProfile.ExcelMatrix.IntraDisplayOrder = displayOrder;
            ExcelComponents.Add(protectionClassProfile);
        }

        

        public void AddConstructionTypeProfile(ISubline item)
        {
            if (!item.IsProperty()) return;

            var displayOrder = ConstructionTypeProfiles.Any()
                ? ConstructionTypeProfiles.Max(x => x.IntraDisplayOrder) + 1
                : 0;

            var constructionTypeProfile = new ConstructionTypeProfile(Id)
            {
                Name = $"{BexConstants.ConstructionTypeName} {item.ShortName}",
                ComponentId = item.Code
            };
            constructionTypeProfile.ExcelMatrix.Add(item);
            constructionTypeProfile.ExcelMatrix.IntraDisplayOrder = displayOrder;
            ExcelComponents.Add(constructionTypeProfile);
        }

        public void AddTotalInsuredValueProfile(ISubline item)
        {
            if (!item.IsProperty()) return;

            var displayOrder = TotalInsuredValueProfiles.Any()
                ? TotalInsuredValueProfiles.Max(x => x.IntraDisplayOrder) + 1
                : 0;

            var totalInsuredValueProfile = new TotalInsuredValueProfile(Id)
            {
                Name = $"{BexConstants.TotalInsuredValueAbbreviatedProfileName} {item.ShortName}",
                ComponentId = item.Code
            };
            totalInsuredValueProfile.ExcelMatrix.Add(item);
            totalInsuredValueProfile.ExcelMatrix.IntraDisplayOrder = displayOrder;
            ExcelComponents.Add(totalInsuredValueProfile);
        }

        public void AddOccupancyTypeProfile(ISubline item)
        {
            var isCommercial = !item.IsPersonal;
            if (!(item.IsProperty() && isCommercial)) return;

            var displayOrder = OccupancyTypeProfiles.Any()
                ? OccupancyTypeProfiles.Max(x => x.IntraDisplayOrder) + 1
                : 0;

            var occupancyTypeProfile = new OccupancyTypeProfile(Id)
            {
                Name = $"{BexConstants.OccupancyTypeName} {item.ShortName}",
                ComponentId = item.Code
            };
            occupancyTypeProfile.ExcelMatrix.Add(item);
            occupancyTypeProfile.ExcelMatrix.IntraDisplayOrder = displayOrder;
            ExcelComponents.Add(occupancyTypeProfile);
        }

        private void AddHazardProfile(ISubline item)
        {
            if (!item.HasHazardProfile) return;

            var displayOrder = HazardProfiles.Any()
                ? HazardProfiles.Max(x => x.IntraDisplayOrder) + 1
                : 0;

            var hazardProfile = new HazardProfile(Id)
            {
                Name = $"{BexConstants.HazardProfileName} {item.ShortNameWithLob}",
                ComponentId = item.Code
            };
            hazardProfile.ExcelMatrix.Add(item);
            hazardProfile.ExcelMatrix.IntraDisplayOrder = displayOrder;
            ExcelComponents.Add(hazardProfile);
        }

        #region collection methods
        public IEnumerator<ISubline> GetEnumerator()
        {
            return _sublines.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(ISubline item)
        {
            Debug.Assert(item != null, $"Can't add a null {BexConstants.SublineName.ToLower()}");

            var row = _sublines.Count(sublineItem => string.Compare(item.ShortNameWithLob, sublineItem.ShortNameWithLob, StringComparison.Ordinal) > 0);
            _sublines.Insert(row, item);

            AddHazardProfile(item);
            AddConstructionTypeProfile(item);
            AddProtectionClassProfile(item);
            AddTotalInsuredValueProfile(item);
            AddOccupancyTypeProfile(item);

            AddWorkersCompStateHazardGroupProfile();
            AddWorkersCompClassCodeProfile();
            AddWorkersCompStateAttachmentProfile();
            AddWorkersCompMinnesotaRetention();

            // ReSharper disable once ExplicitCallerInfoArgument
            NotifyPropertyChanged("SublineViews");
        }
        
        public void Clear()
        {
            _sublines.Clear();
            // ReSharper disable once ExplicitCallerInfoArgument
            NotifyPropertyChanged("SublineViews");
        }

        public bool Contains(ISubline item)
        {
            return _sublines.Contains(item, new SublineComparer());
        }

        public void CopyTo(ISubline[] array, int arrayIndex)
        {
            _sublines.CopyTo(array, arrayIndex);
        }

        public bool Remove(ISubline item)
        {
            var ecs = ExcelComponents.Where(ec => !(ec is HazardProfile) 
                                                  && !(ec is ConstructionTypeProfile)
                                                  && !(ec is ProtectionClassProfile)
                                                  && !(ec is TotalInsuredValueProfile)
                                                  && !(ec is OccupancyTypeProfile)
                                                  && ec.CommonExcelMatrix is MultipleOccurrenceSegmentExcelMatrix).ToList();

            ecs.Where(ec => !(ec is HazardProfile) 
                            && !(ec is ConstructionTypeProfile)
                            && !(ec is ProtectionClassProfile)
                            && !(ec is OccupancyTypeProfile)
                            && !(ec is TotalInsuredValueProfile)
                            ).ForEach(ec =>
            {
                ((MultipleOccurrenceSegmentExcelMatrix)ec.CommonExcelMatrix).Remove(item);
            });

            var hp = HazardProfiles.SingleOrDefault(x => x.ExcelMatrix.Subline.Code == item?.Code);
            if (hp != null)
            {
                ExcelComponents.Remove(hp);
                HazardProfiles.Where(x => x.IntraDisplayOrder > hp.IntraDisplayOrder).ForEach(x => x.ExcelMatrix.IntraDisplayOrder--);
            }

            var ctp = ConstructionTypeProfiles.SingleOrDefault(x => x.ExcelMatrix.Subline.Code == item?.Code);
            if (ctp != null)
            {
                ExcelComponents.Remove(ctp);
                ConstructionTypeProfiles.Where(x => x.IntraDisplayOrder > ctp.IntraDisplayOrder).ForEach(x => x.ExcelMatrix.IntraDisplayOrder--);
            }

            var pcp = ProtectionClassProfiles.SingleOrDefault(x => x.ExcelMatrix.Subline.Code == item?.Code);
            if (pcp != null)
            {
                ExcelComponents.Remove(pcp);
                ProtectionClassProfiles.Where(x => x.IntraDisplayOrder > pcp.IntraDisplayOrder).ForEach(x => x.ExcelMatrix.IntraDisplayOrder--);
            }

            var totalInsuredValueProfile = TotalInsuredValueProfiles.SingleOrDefault(prof => prof.ExcelMatrix.Subline.Code == item?.Code);
            if (totalInsuredValueProfile != null)
            {
                ExcelComponents.Remove(totalInsuredValueProfile);
                TotalInsuredValueProfiles.Where(prof => prof.IntraDisplayOrder > totalInsuredValueProfile.IntraDisplayOrder).ForEach(prof => prof.ExcelMatrix.IntraDisplayOrder--);
            }

            var otp = OccupancyTypeProfiles.SingleOrDefault(prof => prof.ExcelMatrix.Subline.Code == item?.Code);
            if (otp != null)
            {
                ExcelComponents.Remove(otp);
                OccupancyTypeProfiles.Where(prof => prof.IntraDisplayOrder > otp.IntraDisplayOrder).ForEach(prof => prof.ExcelMatrix.IntraDisplayOrder--);
            }


            var result = _sublines.Remove(item);

            // ReSharper disable once ExplicitCallerInfoArgument
            NotifyPropertyChanged("SublineViews");
            return result;
        }

        #endregion
    }
}
