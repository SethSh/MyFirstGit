using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Extensions;
using PionlearClient.Model;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelWorkspaceFolder;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Enums;
using SubmissionCollector.Models.Package.DataComponents;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;
using CedentEditor = SubmissionCollector.View.Editors.CedentEditor;
using UnderwriterEditor = SubmissionCollector.View.Editors.UnderwriterEditor;

// ReSharper disable ExplicitCallerInfoArgument

namespace SubmissionCollector.Models.Package
{
    public sealed class Package : BaseInventoryItem, IPackageInventoryItem, IPackage
    {
        public Package()
        {
            //this is the constructor that will run when restoring from json
            //whoever calls the restoration must RegisterChangeEvents manually
        }

        public Package(PackageModel packageModel)
        {
            IsDirty = true;
            CedentName = packageModel.CedentName;
            CedentId = packageModel.CedentId;
            Guid = packageModel.Guid;
            WorksheetName = "Package";
            Currency = packageModel.Currency;
            Worksheet = (Worksheet) Globals.ThisWorkbook.Worksheets[WorksheetName];
            ExcelMatrices = new List<IExcelMatrix> {new UnderwritingYearExcelMatrix()};
            BexCommunications = new List<BexCommunicationEntry>();

            _segments = new ObservableCollection<ISegment>();
            RegisterChangeEvents();
            Name = packageModel.Name;

            InitializeWorkbook();
        }

        public void RegisterChangeEvents()
        {
            _segments.CollectionChanged += SegmentsCollectionChanged;
            RegisterBaseChangeEvents();
        }

        private string _name;
        private string _cedentId;
        private string _cedentName;
        private string _responsibleUnderwriter;
        private string _currency;
        private string _responsibleUnderwriterId;
        private PackageModel _packageModel;
        private int _underwritingYear;
        private ObservableCollection<ISegment> _segments;
        private string _renewedFromWorkbookName;
        private DateTime? _renewedFromWorkbookDate;

        [PropertyOrder(0)]
        [Category("Attributes")]
        [DisplayName(@"Package Name")]
        [Description("Data Package Name")]
        [Browsable(true)]
        [ReadOnly(false)]
        public new string Name
        {
            get => _name;
            set
            {
                _name = value;
                NotifyPropertyChanged();
                if (Worksheet == null) return;

                using (new ExcelEventDisabler())
                {
                    if (!Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook) Globals.ThisWorkbook.WriteNameToPackageHeader(this);
                }
            }
        }

        [Editor(typeof(CedentEditor), typeof(CedentEditor))]
        [Category("Attributes")]
        [DisplayName(@"Cedent")]
        [Description("Cedent Business Partner ID and Name")]
        [Browsable(true)]
        [ReadOnly(true)]
        public string CedentIdAndName => CedentId.ConnectWithDash(CedentName);

        [Category("Attributes")]
        [DisplayName(@"Cedent Name")]
        [Description("Cedent Business Partner Name")]
        [Browsable(false)]
        [ReadOnly(true)]
        public string CedentName
        {
            get => _cedentName;
            set
            {
                _cedentName = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged("CedentIdAndName");
            }
        }

        [Category("Attributes")]
        [DisplayName(@"Cedent ID")]
        [Browsable(false)]
        [ReadOnly(true)]
        public string CedentId
        {
            get => _cedentId;
            set
            {
                _cedentId = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged("CedentIdAndName");
            }
        }

        [Editor(typeof(UnderwriterEditor), typeof(UnderwriterEditor))]
        [Category("Attributes")]
        [DisplayName(@"Analyst")]
        [Description("Analyst is used to filter packages in BEX.")]
        [Browsable(true)]
        [ReadOnly(true)]
        public string ResponsibleUnderwriter
        {
            get => _responsibleUnderwriter;
            set
            {
                _responsibleUnderwriter = value;
                NotifyPropertyChanged();
            }
        }

        [Category("Server")]
        [DisplayName("ID")]
        [Description("Populated once package is sent to server")]
        [ReadOnly(true)]
        [Browsable(true)]
        public override long? SourceId
        {
            get => base.SourceId;
            set => base.SourceId = value;
        }

        [Browsable(false)]
        public string ResponsibleUnderwriterId
        {
            get => _responsibleUnderwriterId;
            set
            {
                _responsibleUnderwriterId = value;
                NotifyPropertyChanged();
            }
        }

        [Category("Attributes")]
        [DisplayName(@"Currency Type")]
        [Browsable(true)]
        [ReadOnly(true)]
        public string Currency
        {
            get => _currency;
            set
            {
                _currency = value;
                NotifyPropertyChanged();
            }
        }

        [Browsable(false)]
        public ObservableCollection<ISegment> Segments
        {
            get => _segments;
            set
            {
                _segments = value;
                NotifyPropertyChanged();
                NotifyPropertyChanged("SegmentViews");
            }
        }

        [Category("Development")]
        [Browsable(false)]
        [ReadOnly(true)]
        [JsonIgnore]
        public Worksheet Worksheet { get; set; }

        [Category("Development")]
        [Browsable(false)]
        [ReadOnly(true)]
        public string WorksheetName { get; set; }

        [Browsable(false)]
        [JsonIgnore]
        public StringBuilder ExcelValidation { get; set; }

        [JsonIgnore]
        [Browsable(false)]
        public IEnumerable<IInventoryItem> SegmentViews => Segments.OfType<IInventoryItem>().OrderBy(s => s.DisplayOrder).ToList();
        
        [Browsable(false)]
        public List<IExcelMatrix> ExcelMatrices { get; set; }

        [Browsable(false)]
        public UnderwritingYearExcelMatrix UnderwritingYearExcelMatrix => (UnderwritingYearExcelMatrix) ExcelMatrices.Single(
            x => x is UnderwritingYearExcelMatrix);

        [Browsable(false)]
        public IList<BexCommunicationEntry> BexCommunications { get; set; }

        [Browsable(false)]
        public AttachOptions AttachedToRatingAnalysisStatus
        {
            get
            {
                if (!SourceId.HasValue) return AttachOptions.False;

                var bexCommunicationManager = new BexCommunicationManager();
                try
                {
                    var ratingAnalyses = bexCommunicationManager.GetRatingAnalysisIdsUsingThisPackage(SourceId.Value).ToList();
                    if (ratingAnalyses.Any())
                    {
                        return ratingAnalyses.Any(ra =>
                        {
                            Debug.Assert(ra.IsLocked != null, "ra.IsLocked != null");
                            return ra.IsLocked.Value;
                        }) ? AttachOptions.TrueAndLocked : AttachOptions.TrueAndUnlocked;
                    }

                    return AttachOptions.False;
                }
                catch (FileNotFoundException)
                {
                    return AttachOptions.CannotAssess;
                }
            }
        }

        [Category("Renewed From")]
        [DisplayName(@"Workbook Name")]
        [Browsable(true)]
        [ReadOnly(true)]
        [PropertyOrder(0)]
        public string RenewedFromWorkbookName
        {
            get => _renewedFromWorkbookName;
            set
            {
                _renewedFromWorkbookName = value; 
                NotifyPropertyChanged();
            }
        }

        [Category("Renewed From")]
        [DisplayName(@"Time Stamp")]
        [Browsable(true)]
        [ReadOnly(true)]
        [PropertyOrder(1)]
        public DateTime? RenewedFromWorkbookDate
        {
            get => _renewedFromWorkbookDate;
            set
            {
                _renewedFromWorkbookDate = value;
                NotifyPropertyChanged();
            }
        }

        [JsonIgnore]
        [Browsable(false)]
        public string AttachedToRatingAnalysisMessage =>
            $"This {BexConstants.PackageName.ToLower()} " +
            $"is currently {BexConstants.AttachName.ToLower()}ed to a " +
            $"{BexConstants.BexName} {BexConstants.RatingAnalysisName.ToLower()}.";
        
        public StringBuilder Validate()
        {
            var validation = new StringBuilder();

            var dummyAsOfDate = new DateTime(2020, 12, 31);

            var nameValidation = ValidateName();
            if (nameValidation.Length > 0) validation.AppendLine(nameValidation.ToString());

            var underwritingYearValidation = ValidateUnderwritingYearRange();
            if (underwritingYearValidation.Length > 0) validation.AppendLine(underwritingYearValidation.ToString());

            var analystValidation = ValidateAnalyst();
            if (analystValidation.Length > 0) validation.AppendLine(analystValidation.ToString());


            if (validation.Length > 0) return validation;
            _packageModel = new PackageModel
            {
                Name = Name,
                CedentName = CedentName,
                CedentId = CedentId,
                AsOfDate = dummyAsOfDate.ToUniversalTime(),
                AnalystId = ResponsibleUnderwriterId,
                Guid = Guid,
                IsDirty = IsDirty,
                Currency = Currency,
                SourceTimestamp = SourceTimestamp,
                SourceId = SourceId,
                UnderwritingYear = _underwritingYear.ToString(),
                SegmentModels = new List<SegmentModel>()
            };

            return validation;
        }

        public PackageModel CreatePackageModelForQualityControl()
        {
            using (new StatusBarUpdater($"Creating {BexConstants.PackageName.ToLower()} <{Name}> ..."))
            {
                _packageModel = CreatePackageModel();
                
                foreach (var segment in Segments)
                {
                    using (new StatusBarUpdater($"Creating {BexConstants.SegmentName.ToLower()} <{segment.Name}> ..."))
                    {
                        segment.CreateSegmentModelForQualityControl();
                    }
                }
            }
            return _packageModel;
        }

        public PackageModel CreatePackageModel()
        {
            ExcelValidation = new StringBuilder();

            using (new StatusBarUpdater($"Creating {BexConstants.PackageName.ToLower()} <{Name}> ..."))
            {
                var validation = Validate();
                if (validation.Length > 0)
                {
                    ExcelValidation.AppendLine($"Package: {Name}");
                    ExcelValidation.Append(validation);
                    _packageModel = new PackageModel();
                }

                foreach (var segment in Segments)
                {
                    using (new StatusBarUpdater($"Creating {BexConstants.SegmentName.ToLower()} <{segment.Name}> ..."))
                    {
                        var segmentValidation = segment.CreateSegmentModel();
                        if (segmentValidation.Length > 0)
                        {
                            ExcelValidation.AppendLine($"Submission Segment: {segment.Name}");
                            ExcelValidation.Append(segmentValidation);
                            ExcelValidation.AppendLine();
                            ExcelValidation.AppendLine();
                        }
                        else
                        {
                            _packageModel.SegmentModels.Add(segment.SegmentModel);
                        }
                    }
                }
            }
            return _packageModel;
        }

        
        public ISegment GetSegment(int id)
        {
            return Segments.SingleOrDefault(segment => segment.Id == id);
        }

        public ISegment GetSegmentBasedOnDisplayOrder(int displayOrder)
        {
            var segment = Segments.SingleOrDefault(s => s.DisplayOrder == displayOrder);
            return segment;
        }

        public ISegment GetSegmentBasedOnSelected()
        {
            var segment = Segments.SingleOrDefault(s => s.IsSelected);
            return segment;
        }

        internal void DecouplePackage(IWorkbookLogger logger)
        {
            try
            {
                DecouplePackage();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"{BexConstants.DecoupleName.ToStartOfSentence()} {BexConstants.PackageName} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private void DecouplePackage()
        {
            string message; 
            if (WorkbookSaveManager.IsReadOnly)
            {
                message = $"Can't {BexConstants.DecoupleName.ToLower()} a read-only workbook";
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            var sourceId = SourceId;
            if (!sourceId.HasValue)
            {
                message = $"Can't {BexConstants.DecoupleName.ToLower()} this {BexConstants.PackageName.ToLower()}: " +
                          $"not in the {BexConstants.ServerDatabaseName.ToLower()}";
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            message = $"Are you sure you want to {BexConstants.DecoupleName.ToLower()} this {BexConstants.PackageName.ToLower()} " + 
                      $"from the {BexConstants.ServerDatabaseName.ToLower()}?";
            var confirmationResponse = MessageHelper.ShowWithYesNo(message);
            if (confirmationResponse != DialogResult.Yes) return;

            string response;
            using (new CursorToWait())
            {
                DecoupleFromServer();
                response = $"Decoupled <{Name}>";
            }
            BexCommunications.Clear();

            WorkbookSaveManager.Save();

            MessageHelper.Show(response, MessageType.Success);
        }
        
        public void CopyIdsIntoPredecessorIds()
        {
            PredecessorSourceId = SourceId;

            foreach (var segment in Segments)
            {
                segment.PredecessorSourceId = segment.SourceId;
                foreach (var ec in segment.ExcelComponents)
                {
                    ec.PredecessorSourceId  = ec.SourceId;
                }
            }
        }

        internal void DeleteRenewPropertiesAndPredecessorIds(IWorkbookLogger logger)
        {
            try
            {
                DeleteRenewProperties();
                DeletePredecessorIds();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Delete renew from failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public override void DecoupleFromServer()
        {
            base.DecoupleFromServer();
            Segments.ForEach(segment => segment.DecoupleFromServer());
        }

        public void SetAllItemsToNotDirty()
        {
            this.IsDirty = false;
            foreach (var segment in this.Segments)
            {
                segment.IsDirty = false;
                foreach (var excelComponent in segment.ExcelComponents)
                {
                    excelComponent.IsDirty = false;
                }
                segment.SetAllLedgerItemsToNotDirty();
            }
        }

        private void SegmentsCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            RenumberSegmentDisplayOrder();
            NotifyPropertyChanged("SegmentViews");
        }

        private void RenumberSegmentDisplayOrder()
        {
            var i = 0;
            Segments.ForEach(s => s.DisplayOrder = i++);
        }

        private void InitializeWorkbook()
        {
            Worksheet.LockNecessaryCells();
            
            var ws = (Worksheet) Globals.ThisWorkbook.Worksheets[ExcelConstants.SegmentTemplateWorksheetName];
            ws.LockNecessaryCells();

            //hide segment template names
            #if !DEBUG
            foreach (Name nm in Globals.ThisWorkbook.Names)
            {
                if (nm.Name.Contains(ExcelConstants.SegmentTemplateRangeName))
                {
                    nm.Visible = false;
                }
            }
            #endif

            ws.SetVisibleToHidden();
        }

        private StringBuilder ValidateName()
        {
            var validation = new StringBuilder();
            if (Name != BexConstants.DefaultPackageName) return validation;

            validation.AppendLine($"Change {BexConstants.PackageName.ToLower()} from default value <{BexConstants.DefaultPackageName}>");
            return validation;
        }
        private StringBuilder ValidateUnderwritingYearRange()
        {
            var validation = new StringBuilder();

            _underwritingYear = new int();
            var year = UnderwritingYearExcelMatrix.GetInputRange().GetContent()[0, 0];
            if (year == null)
            {
                validation.AppendLine($"Enter a year in {BexConstants.UnderwritingYearName.ToLower()}");
                return validation;
            }

            var yearAsString = year.ToString();
            if (!int.TryParse(yearAsString, out var yearAsInt))
            {
                validation.AppendLine($"Enter a year in {BexConstants.UnderwritingYearName.ToLower()}: <{yearAsString}> is not a year");
            }
            else
            {
                const int underwritingYearTolerance = 10;
                var minimumYear = DateTime.Now.Year - underwritingYearTolerance;
                var maximumYear = DateTime.Now.Year + underwritingYearTolerance;
                if (yearAsInt < minimumYear)
                {
                    validation.AppendLine($"Enter a year greater than {minimumYear} in {BexConstants.UnderwritingYearName.ToLower()}");
                }
                if (yearAsInt > maximumYear)
                {
                    validation.AppendLine($"Enter a year less than {maximumYear} in {BexConstants.UnderwritingYearName.ToLower()}");
                }
            }
            if (validation.Length == 0) _underwritingYear = Convert.ToInt16(yearAsInt);
            return validation;
        }

        private void DeleteRenewProperties()
        {
            RenewedFromWorkbookName = string.Empty;
            RenewedFromWorkbookDate = null;
        }

        private void DeletePredecessorIds()
        {
            PredecessorSourceId = null;

            Segments.ForEach(segment =>
            {
                segment.PredecessorSourceId = null;

                segment.ExcelComponents.ForEach(ec =>
                    {
                        ec.PredecessorSourceId = null;
                    }
                );
            });
        }
        
        private StringBuilder ValidateAnalyst()
        {
            var validation = new StringBuilder();
            if (!string.IsNullOrEmpty(ResponsibleUnderwriterId)) return validation;

            validation.AppendLine($"Enter {BexConstants.AnalystName.ToLower()}");
            return validation;
        }
        
        public void AddItemToBexCommunication(string activity)
        {
            BexCommunications.Add(new ServerCommunication(activity).BexCommunicationEntry);
        }
    }
}


