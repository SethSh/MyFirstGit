using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using PionlearClient.KeyDataFolder;
using Excel = Microsoft.Office.Interop.Excel;
using SubmissionCollector.ActionPanes;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelUtilities.RangeSizeModifier;
using SubmissionCollector.ExcelWorkspaceFolder;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Enums;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.Models.Subline.Property;
using SubmissionCollector.View;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;
// ReSharper disable InconsistentNaming

namespace SubmissionCollector
{
    public partial class ThisWorkbook
    {
        public bool IsCurrentlyOpeningExistingWorkbook { get; set; }
        internal ExcelWorkspace ThisExcelWorkspace { get; private set; }
        internal InventoryActionPane InventoryActionPane { get; private set; }
        public bool LoadButtonDirtyState { get; set; }
        public bool IsCurrentlyUploading { get; set; }

        public void LoadBexCompatibility()
        {
            //this will never be empty
            var workbookVersionAsString = ((object)BexConstants.WorkbookVersionRangeName.GetRange().Value2).ForceContentToString();
            var workbookVersion = Convert.ToDouble(workbookVersionAsString);
            BexCompatibility.GetCompatibility(workbookVersion, ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
        }
        
        public void ShowPackageWorksheet()
        {
            if ((Excel.Worksheet) ActiveSheet != ThisExcelWorkspace.Package.Worksheet) ThisExcelWorkspace.Package.Worksheet.Select();
        }

        public void ShowSegmentWorksheet(ISegment segment)
        {
            if ((Excel.Worksheet) ActiveSheet != segment.WorksheetManager.Worksheet) segment.WorksheetManager.Worksheet.Select();
        }

        public void DeleteFromNameCollection(Excel.Name item)
        {
            item.Delete();
        }

        public void CopyWorksheet(Excel.Worksheet worksheet)
        {
            // ReSharper disable once RedundantCast
            var application = (Microsoft.Office.Interop.Excel.Application)Application;

            var visibility = worksheet.Visible;
            if (visibility != Excel.XlSheetVisibility.xlSheetVisible)
            {
                worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                worksheet.Copy(After: application.ActiveSheet);
                worksheet.Visible = visibility;
            }
            else
            {
                worksheet.Copy(After: application.ActiveSheet);
            }
        }

        public bool IsWorksheetNameAcceptable(string worksheetNameCandidate)
        {
            if (Globals.ThisWorkbook.DoesWorksheetNameExist(worksheetNameCandidate))
                return false;

            if (DoesWorksheetNameExceedMaximumLength(worksheetNameCandidate))
                return false;

            if (DoesWorksheetNameContainInvalidCharacters(worksheetNameCandidate))
                return false;

            return true;
        }

        public string MakeWorksheetNameAcceptable(string worksheetName)
        {
            var acceptableWorksheetName = worksheetName;
            const int upperBound = 100;
            var j = 1;

            while (!IsWorksheetNameAcceptable(acceptableWorksheetName))
            {
                acceptableWorksheetName = $"{worksheetName}.{j++}";
                if (j == upperBound)
                {
                    throw new Exception("Worksheet couldn't be named");
                }
            }
            return acceptableWorksheetName;
        }

        public bool AssertNotEditingCells()
        {
            if (!IsEditing()) return true;

            const string message = "Can't perform this action while editing cells";
            MessageHelper.Show(message, MessageType.Stop);
            return false;
        }

        internal void InsertRowsIntoRange(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new SegmentExcelComponentIdentifier();
                if (!identifier.Validate()) return;

                if (!(identifier.ExcelMatrix is IRangeResizable))
                {
                    MessageHelper.Show(@"The selection must be within a resizable range", MessageType.Stop);
                    return;
                }

                IRangeResizer rangeInserter = null;
                var excelMatrix = identifier.ExcelMatrix;
                var range = identifier.SelectedRange;
                switch (identifier.ExcelMatrix)
                {
                    case PolicyExcelMatrix _:
                        rangeInserter = new PolicyProfileRowInserter();
                        break;
                    case TotalInsuredValueExcelMatrix _:
                        rangeInserter = new TotalInsuredValueProfileRowInserter();
                        break;
                    case WorkersCompClassCodeExcelMatrix _:
                        rangeInserter = new WorkersCompClassCodeProfileRowInserter();
                        break;
                    case PeriodSetExcelMatrix _:
                        rangeInserter = new PeriodSetRowInserter();
                        break;
                    case ExposureSetExcelMatrix _:
                    case AggregateLossSetExcelMatrix _:
                    {
                        excelMatrix = identifier.Segment.PeriodSet.ExcelMatrix;
                        var periodColumn = excelMatrix.RangeName.GetTopLeftCell().Column;
                        rangeInserter = new PeriodSetRowInserter();
                        var offset = identifier.SelectedRange.GetTopLeftCell().Column - periodColumn;
                        range = range.Offset[0, -offset];
                        break;
                    }
                    case ExcelMatrix _:
                        rangeInserter = new IndividualLossSetRowInserter();
                        break;
                    case RateChangeExcelMatrix _:
                        rangeInserter = new RateChangeSetRowInserter();
                        break;
                }

                if (rangeInserter == null)
                {
                    MessageHelper.Show("Matrix doesn't support inserting rows", MessageType.Stop);
                    return;
                }

                if (!rangeInserter.Validate(excelMatrix, range)) return;
                if (!rangeInserter.IsOkToReformat()) return;

                using (new ExcelEventDisabler())
                {
                    rangeInserter.ModifyRange();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Insert rows failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void InsertRowsIntoRangeWithCount(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new SegmentExcelComponentIdentifier();
                if (!identifier.Validate()) return;

                if (!(identifier.ExcelMatrix is IRangeResizable))
                {
                    MessageHelper.Show(@"The selection must be within a resizable range (blue thick border)", MessageType.Stop);
                    return;
                }

                var rowCount = GetRowCount();
                if (!rowCount.HasValue) return;


                IRangeResizer rangeInserter = null;
                var excelMatrix = identifier.ExcelMatrix;
                var range = identifier.SelectedRange.GetTopLeftCell().Resize[rowCount, 1];
                switch (identifier.ExcelMatrix)
                {
                    case PolicyExcelMatrix _:
                        rangeInserter = new PolicyProfileRowInserter();
                        break;
                    case TotalInsuredValueExcelMatrix _:
                        rangeInserter = new TotalInsuredValueProfileRowInserter();
                        break;
                    case WorkersCompClassCodeExcelMatrix _:
                        rangeInserter = new WorkersCompClassCodeProfileRowInserter();
                        break;
                    case PeriodSetExcelMatrix _:
                        rangeInserter = new PeriodSetRowInserter();
                        break;

                    case ExposureSetExcelMatrix _:
                    case AggregateLossSetExcelMatrix _:
                    {
                        rangeInserter = new PeriodSetRowInserter();
                        excelMatrix = identifier.Segment.PeriodSet.ExcelMatrix;
                        var periodColumn = excelMatrix.RangeName.GetTopLeftCell().Column;
                        var offset = identifier.SelectedRange.GetTopLeftCell().Column - periodColumn;
                        range = range.Offset[0, -offset];
                        break;
                    }

                    case ExcelMatrix _:
                        rangeInserter = new IndividualLossSetRowInserter();
                        break;
                    case RateChangeExcelMatrix _:
                        rangeInserter = new RateChangeSetRowInserter(); 
                        break;
                }

                if (rangeInserter == null)
                {
                    MessageHelper.Show("Matrix doesn't support inserts rows with count", MessageType.Stop);
                    return;
                }
                
                if (!rangeInserter.Validate(excelMatrix, range)) return;
                if (!rangeInserter.IsOkToReformat()) return;

                using (new ExcelEventDisabler())
                {
                    rangeInserter.ModifyRange();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Insert rows failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void DeleteRowsFromRange(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new SegmentExcelComponentIdentifier();
                if (!identifier.Validate()) return;

                if (!(identifier.ExcelMatrix is IRangeResizable))
                {
                    MessageHelper.Show(@"The selection must be within a resizable range", MessageType.Stop);
                    return;
                }

                IRangeResizer rangeDeleter = null;
                ISegmentExcelMatrix excelMatrix = null;
                Excel.Range range = null;
                if (identifier.ExcelMatrix is PolicyExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new PolicyProfileRowDeleter();
                    range = identifier.SelectedRange;
                }
                else if (identifier.ExcelMatrix is TotalInsuredValueExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new TotalInsuredValueProfileRowDeleter();
                    range = identifier.SelectedRange;
                }
                else if (identifier.ExcelMatrix is WorkersCompClassCodeExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new WorkersCompClassCodeProfileRowDeleter();
                    range = identifier.SelectedRange;
                }
                else if (identifier.ExcelMatrix is PeriodSetExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new PeriodSetRowDeleter();
                    range = identifier.SelectedRange;
                }
                else if (identifier.ExcelMatrix is ExposureSetExcelMatrix ||
                         identifier.ExcelMatrix is AggregateLossSetExcelMatrix)
                {
                    excelMatrix = identifier.Segment.PeriodSet.ExcelMatrix;
                    var periodColumn = excelMatrix.RangeName.GetTopLeftCell().Column;
                    rangeDeleter = new PeriodSetRowDeleter();
                    var offset = identifier.SelectedRange.GetTopLeftCell().Column - periodColumn;
                    range = identifier.SelectedRange.Offset[0, -offset];
                }
                else if (identifier.ExcelMatrix is ExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new IndividualLossSetRowDeleter();
                    range = identifier.SelectedRange;
                }
                else if (identifier.ExcelMatrix is RateChangeExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new RateChangeSetRowDeleter();
                    range = identifier.SelectedRange;
                }

                if (rangeDeleter == null)
                {
                    MessageHelper.Show("Matrix doesn't support deleting rows", MessageType.Stop);
                    return;
                }

                if (!rangeDeleter.Validate(excelMatrix, range)) return;
                if (!rangeDeleter.IsOkToReformat()) return;

                rangeDeleter.ModifyRange();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Delete rows failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void InsertColumnsIntoRange(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new SegmentExcelComponentIdentifier();
                if (!identifier.Validate()) return;

                if (!(identifier.ExcelMatrix is IRangeResizable))
                {
                    MessageHelper.Show(@"The selection must be within a resizable range", MessageType.Stop);
                    return;
                }

                IRangeResizer rangeInserter = null;
                ISegmentExcelMatrix excelMatrix = null;
                Excel.Range range = null;
                if (identifier.ExcelMatrix is PolicyExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeInserter = new PolicyProfileColumnInserter();
                    range = identifier.SelectedRange;
                }
                else if (identifier.ExcelMatrix is WorkersCompStateAttachmentExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeInserter = new WorkersCompStateAttachmentColumnInserter();
                    range = identifier.SelectedRange;
                }

                if (rangeInserter == null)
                {
                    MessageHelper.Show("Matrix doesn't support inserting columns", MessageType.Stop);
                    return;
                }
                
                if (!rangeInserter.Validate(excelMatrix, range)) return;
                if (!rangeInserter.IsOkToReformat()) return;

                using (new ExcelEventDisabler())
                {
                    rangeInserter.ModifyRange();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Insert columns failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void DeleteColumnsFromRange(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new SegmentExcelComponentIdentifier();
                if (!identifier.Validate()) return;

                if (!(identifier.ExcelMatrix is IRangeResizable))
                {
                    MessageHelper.Show(@"The selection must be within a resizable range", MessageType.Stop);
                    return;
                }

                IRangeResizer rangeDeleter = null;
                ISegmentExcelMatrix excelMatrix = null;
                Excel.Range range = null;

                if (identifier.ExcelMatrix is PolicyExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new PolicyProfileColumnDeleter();
                    range = identifier.SelectedRange;
                }
                else if (identifier.ExcelMatrix is WorkersCompStateAttachmentExcelMatrix)
                {
                    excelMatrix = identifier.ExcelMatrix;
                    rangeDeleter = new WorkersCompStateAttachmentColumnDeleter();
                    range = identifier.SelectedRange;
                }

                if (rangeDeleter == null)
                {
                    MessageHelper.Show("Matrix doesn't support deleting columns", MessageType.Stop);
                    return;
                }

                if (!rangeDeleter.Validate(excelMatrix, range)) return;
                if (!rangeDeleter.IsOkToReformat()) return;

                rangeDeleter.ModifyRange();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Delete columns failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void TogglePane(Excel.Application application, RibbonToggleButton button,  IWorkbookLogger logger)
        {
            try
            {
                application.DisplayDocumentActionTaskPane = button.Checked;
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Toggle button failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public void AddInventoryToActionsPane()
        {
            ActionsPane.Controls.Add(InventoryActionPane);
            InventoryActionPane.Dock = DockStyle.Fill;

            SetPaneDisplay();
        }

        public void SetPaneDisplay(double width, PanePositionType position)
        {
            const string taskPaneName = "Task Pane";

            // ReSharper disable once RedundantCast
            var application = (Microsoft.Office.Interop.Excel.Application)Application;

            var taskPane = application.CommandBars[taskPaneName];
            taskPane.Position = position == PanePositionType.Left ? MsoBarPosition.msoBarLeft : MsoBarPosition.msoBarRight;
            taskPane.Width = Convert.ToInt32(Math.Round(application.Width * width, MidpointRounding.AwayFromZero));
        }

        public Excel.Range GetSelectedRange()
        {
            // ReSharper disable once RedundantCast
            var application = (Microsoft.Office.Interop.Excel.Application)Application; 
            return application.Selection as Excel.Range;
        }

        public Excel.Worksheet GetSelectedWorksheet()
        {
            return ActiveSheet as Excel.Worksheet;
        }

        public bool HasWorkbookBeenSaved()
        {
            return !string.IsNullOrEmpty(Path);
        }

        public void RenameSelectedWorksheet()
        {
            var selectedSheet = Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet;
            RenameWorksheet(selectedSheet);
        }

        public void MoveSelectedWorksheetRight()
        {
            var selectedSheet = Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet;
            MoveWorksheetRight(selectedSheet);
        }

        public void MoveSelectedWorksheetLeft()
        {
            var selectedSheet = Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet;
            MoveWorksheetLeft(selectedSheet);
        }

        internal bool IsWorksheetOwnedByUser(Excel.Worksheet ws)
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            return package.Worksheet != ws && package.Segments.All(x => ws != x.WorksheetManager.Worksheet);
        }

        internal bool IsSelectedWorksheetOwnedByUser()
        {
            return !(Globals.ThisWorkbook.ActiveSheet is Excel.Worksheet selectedSheet) || IsWorksheetOwnedByUser(selectedSheet);
        }

        internal void WriteSegmentNameToHeader(ISegment segment)
        {
            var range = segment.HeaderRangeName.GetRange();
            range.GetBottomRightCell().Value2 = segment.Name;
        }

        internal void WriteNameToPackageHeader(Package package)
        {
            var name = package.Name;
            var packageRange = package.Worksheet.Range[$"submission.{ExcelConstants.HeaderRangeName}"].Resize[1, 1].Offset[0, 1];

            packageRange.Value2 = name;
            package.Segments.ForEach(x => x.HeaderRangeName.GetRange().GetTopRightCell().Value2 = name);
        }

        internal bool DoesWorksheetNameExceedMaximumLength(string worksheetName)
        {
            return worksheetName.Length > ExcelConstants.WorksheetNameMaximumCharacters;
        }

        internal bool DoesWorksheetNameContainInvalidCharacters(string worksheetName)
        {
            var invalidCharacters = new[] {':', '\\', '/', '?', '*', '[', ']'};
            return invalidCharacters.Any(worksheetName.Contains);
        }

        internal bool DoesWorksheetNameExist(string worksheetName)
        {
            return Worksheets.Cast<Excel.Worksheet>().Any(item => item.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase));
        }

        internal void ReformatDataComponent(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new SegmentExcelComponentIdentifier();
                if (!identifier.Validate()) return;

                using (new ExcelScreenUpdateDisabler())
                {
                    using (new ExcelEventDisabler())
                    {
                        identifier.ExcelMatrix.Reformat();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Revert format failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void DeleteWorksheet(Excel.Worksheet worksheet)
        {
            using (new WorkbookUnprotector())
            {
                using (new ExcelEventDisabler())
                {
                    using (new ExcelAlertDisabler())
                    {
                        worksheet.Delete();
                        WorkbookExtensions.RemoveErrorRanges();
                    }
                }
            }
        }

        internal void RenameWorksheet(Excel.Worksheet worksheet)
        {
            var name = MessageHelper.ShowInputBox("Enter new worksheet name", worksheet.Name);
            if (name == worksheet.Name) return;

            try
            {
                using (new WorkbookUnprotector())
                {
                    worksheet.Name = name;
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Show(ex.Message, MessageType.Stop);
            }

        }

        internal void DeleteSelectedWorksheet()
        {
            //already checked to ensure it's a worksheet
            var worksheet = (Excel.Worksheet)Globals.ThisWorkbook.ActiveSheet;
            DeleteWorksheet(worksheet);
        }
        
        private void SetPaneDisplay()
        {
            var paneWidthFactor = UserPreferences.ReadFromFile().PaneWidthFactor;
            var position = UserPreferences.ReadFromFile().PanePosition;
            SetPaneDisplay(paneWidthFactor, position);
        }

        private void ThisWorkbook_Startup(object sender, EventArgs e)
        {
            var stackTraceLogger = new StackTraceLogger();
            try
            {
                System.Net.ServicePointManager.ServerCertificateValidationCallback = (caller, cert, chain, errors) => true;

                DownloadReferenceData(); 
                CreateUserPreferences(); 
                SetWorkbookVersion();

                Task.Factory.StartNew(LoadBexCompatibility);
                SetExcelEvents();
                
                var bexCustomXmlPart = CustomXmlPartManager.GetBexCustomXmlPart();
                IsCurrentlyOpeningExistingWorkbook = bexCustomXmlPart != null;

                ThisExcelWorkspace = IsCurrentlyOpeningExistingWorkbook ? new ExcelWorkspace(bexCustomXmlPart) : new ExcelWorkspace();

                ProtectWorksheets();
                ProtectTemplateWorksheet();
                
                InventoryActionPane = new InventoryActionPane();
                AddInventoryToActionsPane();
                Globals.Ribbons.SubmissionRibbon.PaneToggleButton.Checked = true;

                var excelSheetActivateEventManager = new ExcelSheetActivateEventManager();
                excelSheetActivateEventManager.MonitorSheetChange(GetSelectedWorksheet(), ThisExcelWorkspace.Package);

                WriteInUnderwritingYear();
                WriteInUnderwriter();

                var isAdmin = UserPreferences.ReadFromFile().IsAdmin;
                Globals.Ribbons.SubmissionRibbon.AdminGroup.Visible = isAdmin;

                RangeNameDisplayer.SetVisibility(false, new StackTraceLogger());
                RetroFitPropertyRanges();

                DisplayWarnings();
            }
            catch (Exception ex)
            {
                const string errorMessage = "Workbook start up failed";
                stackTraceLogger.WriteNew(ex);
                MessageHelper.Show(errorMessage, MessageType.Stop);
            }
            finally
            {
                IsCurrentlyOpeningExistingWorkbook = false;
            }
        }

        private void RetroFitPropertyRanges()
        {
            var package = ThisExcelWorkspace.Package;
            foreach (var segment in package.Segments)
            {
                if (!segment.IsProperty) continue;

                foreach (var profile in segment.TotalInsuredValueProfiles)
                {
                    var isCommercial = profile.ExcelMatrix.ContainsCommercial();
                    var sublineCode = isCommercial
                        ? BexConstants.CommercialPropertySublineCode
                        : BexConstants.GeneralLiabilityHomeownersPropertySublineCode;

                    var sublineAbbreviation = isCommercial ? "Comm" : "Pers";
                    var sublineLabel = isCommercial
                        ? new CommercialSubline().ShortNameWithLob
                        : new HomeownersPropertySubline().ShortNameWithLob;

                    var oldMatchingIndex = GetMatchingIndex(segment, ExcelConstants.TotalInsuredValueProfileRangeName, sublineLabel);
                    if (!oldMatchingIndex.HasValue) continue;
                    
                    var oldRangeName = TotalInsuredValueExcelMatrix.GetRangeName(segment.Id, oldMatchingIndex.Value);
                    var newRangeName = TotalInsuredValueExcelMatrix.GetRangeName(segment.Id, sublineCode);
                    oldRangeName.RenameRange(newRangeName);

                    var oldBasisRangeName = TotalInsuredValueExcelMatrix.GetBasisRangeName(segment.Id, oldMatchingIndex.Value);
                    var newBasisRangeName = TotalInsuredValueExcelMatrix.GetBasisRangeName(segment.Id, sublineCode);
                    oldBasisRangeName.RenameRange(newBasisRangeName);

                    var oldHeaderRangeName = TotalInsuredValueExcelMatrix.GetHeaderRangeName(segment.Id, oldMatchingIndex.Value);
                    var newHeaderRangeName = TotalInsuredValueExcelMatrix.GetHeaderRangeName(segment.Id, sublineCode);

                    oldHeaderRangeName.RenameRange(newHeaderRangeName);

                    profile.Name = $"{BexConstants.TotalInsuredValueAbbreviatedProfileName} {sublineAbbreviation}";
                    profile.ComponentId = sublineCode;

                    using (new ExcelEventDisabler())
                    {
                        profile.ExcelMatrix.GetHeaderRange().GetTopLeftCell().Value = profile.Name;

                        var oldSublinesRangeName = TotalInsuredValueExcelMatrix.GetSublinesRangeName(segment.Id, oldMatchingIndex.Value);
                        oldSublinesRangeName.GetRange().Clear();
                        oldSublinesRangeName.DeleteRangeName();

                        var oldSublinesHeaderRangeName = TotalInsuredValueExcelMatrix.GetSublinesHeaderRangeName(segment.Id, oldMatchingIndex.Value);
                        oldSublinesHeaderRangeName.GetRange().Clear();
                        oldSublinesHeaderRangeName.DeleteRangeName();
                    }
                }

                foreach (var profile in segment.OccupancyTypeProfiles)
                {
                    const int sublineCode = BexConstants.CommercialPropertySublineCode;
                    const string sublineAbbreviation = "Comm";
                    var sublineLabel = new CommercialSubline().ShortNameWithLob;

                    var oldMatchingIndex = GetMatchingIndex(segment, ExcelConstants.OccupancyTypeProfileRangeName, sublineLabel);
                    if (!oldMatchingIndex.HasValue) continue;
                    
                    var oldRangeName = OccupancyTypeExcelMatrix.GetRangeName(segment.Id, oldMatchingIndex.Value);
                    var newRangeName = OccupancyTypeExcelMatrix.GetRangeName(segment.Id, sublineCode);
                    oldRangeName.RenameRange(newRangeName);

                    var oldBasisRangeName = OccupancyTypeExcelMatrix.GetBasisRangeName(segment.Id, oldMatchingIndex.Value);
                    var newBasisRangeName = OccupancyTypeExcelMatrix.GetBasisRangeName(segment.Id, sublineCode);
                    oldBasisRangeName.RenameRange(newBasisRangeName);

                    var oldHeaderRangeName = OccupancyTypeExcelMatrix.GetHeaderRangeName(segment.Id, oldMatchingIndex.Value);
                    var newHeaderRangeName = OccupancyTypeExcelMatrix.GetHeaderRangeName(segment.Id, sublineCode);
                    oldHeaderRangeName.RenameRange(newHeaderRangeName);

                    profile.Name = $"{BexConstants.OccupancyTypeName} {sublineAbbreviation}";
                    profile.ComponentId = sublineCode;

                    using (new ExcelEventDisabler())
                    {
                        profile.ExcelMatrix.GetHeaderRange().GetTopLeftCell().Value = profile.Name;

                        var oldSublinesRangeName = OccupancyTypeExcelMatrix.GetSublinesRangeName(segment.Id, oldMatchingIndex.Value);
                        oldSublinesRangeName.GetRange().Clear();
                        oldSublinesRangeName.DeleteRangeName();

                        var oldSublinesHeaderRangeName = OccupancyTypeExcelMatrix.GetSublinesHeaderRangeName(segment.Id, oldMatchingIndex.Value);
                        oldSublinesHeaderRangeName.GetRange().Clear();
                        oldSublinesHeaderRangeName.DeleteRangeName();
                    }
                }
            }
        }

        private static int? GetMatchingIndex(ISegment segment, string profileString, string sublineLabel)
        {
            //identify an old range name without having the component id
            //new ranges will not have any sublines range
            var matchingIndex = new int?();
            for (var index = 0; index < BexConstants.MaximumNumberOfDataComponents; index++)
            {
                var oldSublinesRangeName = $"segment{segment.Id}.{profileString}{index}.sublines";
                var oldSublinesRange = oldSublinesRangeName.GetRangeOrDefault();
                if (oldSublinesRange == null) continue;

                var rangeContent = oldSublinesRange.GetContent().ForceContentToStrings();
                for (var row = 0; row < rangeContent.GetLength(0); row++)
                {
                    if (rangeContent[row, 0] == sublineLabel)
                    {
                        matchingIndex = index;
                        break;
                    }
                }

                if (matchingIndex.HasValue) break;
            }

            return matchingIndex;
        }

        private void DisplayWarnings()
        {
            var package = ThisExcelWorkspace.Package;
            if (!package.SourceId.HasValue) return;

            var whileClause = $"While data {BexConstants.UpdateName.ToLower()}s are permitted, " +
                              $"structural {BexConstants.UpdateName.ToLower()}s are blocked.";

            var lockedClause =
                $"{BexConstants.UploadName.ToStartOfSentence()}s will be blocked because this {BexConstants.PackageName.ToLower()} is " +
                $"{BexConstants.AttachName.ToLower()}ed to a locked {BexConstants.RatingAnalysisName.ToLower()}.";

            var sb = new StringBuilder();
            switch (package.AttachedToRatingAnalysisStatus)
            {
                case AttachOptions.TrueAndLocked:
                    sb.Append(lockedClause);
                    MessageHelper.Show(sb.ToString(), MessageType.Warning);
                    break;

                case AttachOptions.TrueAndUnlocked:
                    sb.Append(package.AttachedToRatingAnalysisMessage + Environment.NewLine + whileClause);
                    MessageHelper.Show(sb.ToString(), MessageType.Warning);
                    break;

                case AttachOptions.CannotAssess:
                    sb.Append($"Can't connect to the {BexConstants.BexName} API." + Environment.NewLine + whileClause);
                    MessageHelper.Show(sb.ToString(), MessageType.Warning);
                    break;
                
                case AttachOptions.False:
                    break;
                
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static void SetWorkbookVersion()
        {
            if (BexConstants.WorkbookVersionRangeName.ExistsInWorkbook()) return;

            var versionRange = BexConstants.SubmissionHeaderRangeName.GetRange().GetBottomRightCell();
            versionRange.SetInvisibleRangeName( BexConstants.WorkbookVersionRangeName);

            if (versionRange.Value2 == null) versionRange.Value2 = BexConstants.WorkbookVersion.ToString("N2");
        }

        private bool WorkbookTemplateIsOpening()
        {
            return !HasWorkbookBeenSaved();
        }

        private void WriteInUnderwriter()
        {
            if (!WorkbookTemplateIsOpening()) return;

            var up = UserPreferences.ReadFromFile();
            if (!up.UseMeAsUnderwriter) return;

            var name = Environment.UserName;
            var underwriters = UnderwritersFromKeyData.UnderwriterReferenceData;

            var underwriter = underwriters.FirstOrDefault(und => und.Code == name);
            if (underwriter == null) return;
            
            var package = ThisExcelWorkspace.Package;
            package.ResponsibleUnderwriterId = underwriter.Code;
            package.ResponsibleUnderwriter = underwriter.Name;
        }

        private void WriteInUnderwritingYear()
        {
            if (!WorkbookTemplateIsOpening()) return;
            
            var up = UserPreferences.ReadFromFile();
            if (up.UnderwritingYearType == UnderwritingYearType.None) return;

            var range = ThisExcelWorkspace.Package.UnderwritingYearExcelMatrix.GetInputRange().GetTopLeftCell();
            var content = range.GetContent().ForceContentToStrings()[0, 0];
            if (content != null) return;

            switch (up.UnderwritingYearType)
            {
                case UnderwritingYearType.CurrentYear:
                    range.Value2 = DateTime.Now.Year;
                    return;

                case UnderwritingYearType.NextYear:
                    range.Value2 = DateTime.Now.Year + 1;
                    return;

                case UnderwritingYearType.None:
                    break;

                default:
                    throw new ArgumentOutOfRangeException($"{BexConstants.UnderwritingYearName.ToStartOfSentence()} default can't be set");

            }
        }
        
        // ReSharper disable once InconsistentNaming
        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            if (Wb.Name != Name) return;

            var preferences = UserPreferences.ReadFromFile();
            if (!preferences.SubmissionRibbonSelectionChoice) return;

            Globals.Ribbons.SubmissionRibbon.SelectSubmissionRibbonWithWait();
        }

        private static void CreateUserPreferences()
        {
            if (UserPreferences.FileExists) return;

            var up = new UserPreferences();
            up.CreateNew();
            up.WriteToFile();
        }

        private void ThisWorkbook_SheetChange(object Sh, Excel.Range Target)
        {
            if (ThisExcelWorkspace == null) return;

            var excelSheetChangeEventManager = new ExcelSheetChangeEventManager();
            excelSheetChangeEventManager.MonitorSheetChange(Sh as Excel.Worksheet, Target);
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            if (Wb.Name != Name) return;
            if (Saved) return;

            var up = UserPreferences.ReadFromFile();
            if (!up.CheckSynchronization) return;

            if (LoadButtonDirtyState) return;

            var synchronizationRenderer = new SynchronizationManager();
            var isSynchronized = synchronizationRenderer.RenderHeadless();
            if (isSynchronized) return;

            Globals.Ribbons.SubmissionRibbon.SetSynchronizationButtonImage(true);
            var message = $"The workbook content and the {BexConstants.ServerDatabaseName.ToLower()} content are no longer synchronized.  " +
                          $"(The {BexConstants.UploadName.ToLower()} ribbon button has changed from green to red.)  " +
                          "Are you sure you want to close?";
            var response = MessageHelper.ShowWithYesNo(message);
            if (response == DialogResult.No) Cancel = true;
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            if (Wb.Name != Name) return;

            var up = UserPreferences.ReadFromFile();
            if (up.CheckSynchronization)
            {
                if (!IsCurrentlyUploading)
                {
                    var appearsSynchronized = !LoadButtonDirtyState;
                    if (appearsSynchronized)
                    {
                        var synchronizationRenderer = new SynchronizationManager();
                        var isSynchronized = synchronizationRenderer.RenderHeadless();
                        if (!isSynchronized)
                        {
                            Globals.Ribbons.SubmissionRibbon.SetSynchronizationButtonImage(true);
                            MessageHelper.Show("The system changed the upload ribbon button from green to red", MessageType.Information);
                        }
                    }
                }
            }

            var manager = new CustomXmlPartManager(new StackTraceLogger());
            manager.WriteJsonToCustomXmlPart(ThisExcelWorkspace.Package, this);
        }

        // ReSharper disable once InconsistentNaming
        private void ThisWorkbook_SheetActivate(object Sh)
        {
            if (ThisExcelWorkspace == null) return;

            var sheetActivateEventManager = new ExcelSheetActivateEventManager();
            sheetActivateEventManager.MonitorSheetChange(Sh as Excel.Worksheet, ThisExcelWorkspace.Package);
        }

        private void ActionsPane_VisibleChanged(object sender, EventArgs e)
        {
            //user closing panel doesn't raise event
            //Determining When the Actions Pane is Closed
            //There is no event that is raised when the actions pane is closed. Although the ActionsPane class has a VisibleChanged event, this event is not raised when the end user closes the actions pane. 
            //Instead, this event is raised when the controls on the actions pane are hidden by calling the Hide method or by setting the Visible property to false.
            //https://msdn.microsoft.com/en-us/library/7we49he1.aspx
        }

        
        // ReSharper disable once MemberCanBeMadeStatic.Local
        private void ThisWorkbook_Shutdown(object sender, EventArgs e)
        {
            Dispatcher.CurrentDispatcher.InvokeShutdown();
        }

        private void ProtectWorksheets()
        {
            //because only protecting from ui and not code, this needs to get reset on open
            ThisExcelWorkspace.Package.Worksheet.ProtectInterface();
            foreach (var item in ThisExcelWorkspace.Package.Segments)
            {
                item.WorksheetManager.Worksheet.ProtectInterface();
            }
        }

        private void ProtectTemplateWorksheet()
        {
            // ReSharper disable once RedundantCast
            var sheets = (Excel.Sheets)Worksheets;
            var ws = (Excel.Worksheet) sheets[ExcelConstants.SegmentTemplateWorksheetName];
            ws.ProtectInterface();
        }

        
        private void MoveWorksheetRight(Excel.Worksheet selectedSheet)
        {
            // ReSharper disable once RedundantCast
            var sheets = (Excel.Sheets)Worksheets; 
            
            var index = selectedSheet.Index;
            var worksheetCount = sheets.Count;
            if (index == worksheetCount)
            {
                MessageHelper.Show("Unable to move any farther right", MessageType.Stop);
                return;
            }

            using (new WorkbookUnprotector())
            {
                selectedSheet.Move(After: sheets[index + 1]);
            }
        }

        private void MoveWorksheetLeft(Excel.Worksheet selectedSheet)
        {
            var userWorksheetCount = Worksheets.Cast<Excel.Worksheet>()
                .Count(ws => IsWorksheetOwnedByUser(ws) && ws.Visible == Excel.XlSheetVisibility.xlSheetVisible);

            // ReSharper disable once RedundantCast
            var sheets = (Excel.Sheets)Worksheets;
            
            var index = selectedSheet.Index;
            var worksheetCount = sheets.Count;
            var minIndex = worksheetCount - userWorksheetCount + 1;

            if (index == minIndex)
            {
                MessageHelper.Show("Unable to move any farther left", MessageType.Stop);
                return;
            }

            using (new WorkbookUnprotector())
            {
                selectedSheet.Move(Before: sheets[index - 1]);
            }
        }

        private bool IsEditing()
        {
            // ReSharper disable once RedundantCast
            var application = (Excel.Application) Application;
            
            if(application.Interactive == false)
            {
                return false;
            }

            try
            {
                application.Interactive = false;
                application.Interactive = true;
            }
            catch (Exception)
            {
                return true;
            }

            return false;
        }

        private void SetExcelEvents()
        {
            // ReSharper disable once RedundantCast
            var application = (Excel.Application)Application; 
            
            application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            application.WorkbookActivate += Application_WorkbookActivate;
            ActionsPane.VisibleChanged += ActionsPane_VisibleChanged;
            SheetActivate += ThisWorkbook_SheetActivate;
            SheetChange += ThisWorkbook_SheetChange;
        }

        private static void DownloadReferenceData()
        {
            ReferenceDataManager.Download();
        }

        private static int? GetRowCount()
        {
            const int lowerBound = 1;
            const int upperBound = 10000;
            var bound = $"{lowerBound:N0} and {upperBound:N0}";

            var defaultInsertRowCount = UserPreferences.ReadFromFile().InsertRowCount;
            var rowCountAsString = Microsoft.VisualBasic.Interaction.InputBox($"Enter number of rows (between {bound}) to insert " +
                                                                              "into next-to-last input range", "Insert Rows", 
                                                                              defaultInsertRowCount.ToString());
            if (int.TryParse(rowCountAsString, out var rowCount))
            {
                if (rowCount >= 1 && rowCount <= upperBound) return rowCount;

                var message = $"Row count <{rowCount:N0}> doesn't fall with between {bound}";
                MessageBox.Show(message);
                return new int?();
            }

            var message2 = $"Row count not recognized as a number between {bound}";
            MessageBox.Show(message2);
            return new int?();
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            // ReSharper disable once RedundantDelegateCreation
            Startup += new EventHandler(ThisWorkbook_Startup);
            // ReSharper disable once RedundantDelegateCreation
            Shutdown += new EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
