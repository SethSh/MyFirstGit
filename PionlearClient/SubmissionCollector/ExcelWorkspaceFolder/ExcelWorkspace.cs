using System;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using PionlearClient.Model;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    public class ExcelWorkspace
    {
        public ExcelWorkspace()
        {
            var packageModel = new PackageModel();
            Package = new Package(packageModel);
        }
    
        public ExcelWorkspace(CustomXMLPart bexCustomXmlPart)
        {
            var manager = new CustomXmlPartManager(new StackTraceLogger());
            Package = manager.RestoreFromCustomXmlPart(bexCustomXmlPart);
            manager.RestoreWorksheetsFromJson(Package);

            Package.RegisterChangeEvents();
            Package.Segments.ForEach(x => x.RegisterBaseChangeEvents());
        }

        public Package Package;

        internal void DuplicateSelectedSegment(IWorkbookLogger logger)
        {
            try
            {
                var validator = new SegmentWorksheetValidator();
                if (!validator.Validate()) return;
                if (!validator.Segment.IsNameAcceptableLength()) return;

                if (!validator.Segment.IsSelected)
                {
                    var message = $"Duplicating requires first selecting a {BexConstants.SegmentName.ToLower()} node " +
                                  $"in the {BexConstants.InventoryTreeName.ToLower()}.";
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }

                var newSegment = validator.Segment.Duplicate();
                Package.Segments.Add(newSegment);

                var summaryBuilder = new ProspectiveExposureSummaryBuilder();
                summaryBuilder.Build();

                ExcelSheetActivateEventManager.RefreshUmbrellaWizardButton(newSegment);
                Globals.Ribbons.SubmissionRibbon.SetSynchronizationButtonImage(isDirty: true);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Duplicate {BexConstants.SegmentName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void DeleteSegment(IWorkbookLogger logger)
        {
            try
            {
                var isSynchronizedWithServer = !Globals.ThisWorkbook.LoadButtonDirtyState;
                var segment = Package.GetSegmentBasedOnSelected();  
                if (segment == null)
                {
                    const string message = "Segment not selected in the inventory tree";
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }

                if (!segment.IsStructureModifiable)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine(Package.AttachedToRatingAnalysisMessage);
                    sb.AppendLine();
                    sb.AppendLine($"Deleting {BexConstants.SegmentName.ToLower()} <{segment.Name}> is blocked.");
                    MessageHelper.Show(sb.ToString(), MessageType.Stop);
                    return;
                }

                using (new StatusBarUpdater("Deleting submission segment ..."))
                {
                    var confirmationResponse = MessageHelper.ShowWithYesNo($"Are you sure you want to delete {BexConstants.SegmentName.ToLower()} <{segment.Name}>?");
                    if (confirmationResponse != DialogResult.Yes) return;

                    var segmentOnServer = segment.SourceId.HasValue;
                    DeleteSegment(segment);

                    if (isSynchronizedWithServer && segmentOnServer)
                    {
                        Globals.Ribbons.SubmissionRibbon.SetSynchronizationButtonImage(isDirty: true);
                    }

                    var sheetActivateEventManager = new ExcelSheetActivateEventManager();
                    sheetActivateEventManager.MonitorSheetChange(Globals.ThisWorkbook.GetSelectedWorksheet(), Package);
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Delete {BexConstants.SegmentName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public string GetVersionStatus()
        {
            try
            {
                Globals.ThisWorkbook.LoadBexCompatibility();
                var message = BexCompatibility.IsCompatible ? "is" : "isn't";
                return $"Workbook version {BexConstants.WorkbookVersion} {message} compatible with {BexConstants.BexName}";
            }
            catch (Exception ex)
            {
                return $"Unable to verify workbook version compatibility:\n\n{ex.Message}.";
            }
        }

        internal void AddSegment(IWorkbookLogger logger)
        {
            try
            {
                ISegment segment = new Segment(isNotCalledFromJson: true);

                var sublineSelectorResponse = SublineWizardManager.LaunchSublineSelectorWizard(segment);
                if (sublineSelectorResponse.FormResponse == FormResponse.Cancel) return;

                SublineWizardManager.MapSublineWizardViewModel(segment, sublineSelectorResponse.ViewModel);
                Package.Segments.Add(segment);
                using (new ExcelEventDisabler())
                {
                    segment.WorksheetManager.CreateWorksheet();
                    segment.IsSelected = true;
                }

                var summaryBuilder = new ProspectiveExposureSummaryBuilder();
                summaryBuilder.Build();

                ExcelSheetActivateEventManager.RefreshUmbrellaWizardButton(segment);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Add {BexConstants.SegmentName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
        
        internal void ToggleEstimate(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new ExcelComponentIdentifier();
                if (!identifier.Validate()) return;

                using (new ExcelScreenUpdateDisabler())
                {
                    identifier.ExcelMatrix.ToggleEstimate();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Toggle estimate failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void GetBuildVersion(IWorkbookLogger logger)
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                var buildDate = new DateTime(2000, 1, 1).AddDays(version.Build).AddSeconds(version.Revision * 2);

                var sb = new StringBuilder();
                sb.AppendLine($"Version: {version}");
                sb.AppendLine($"Date Built: {buildDate}");
            
                MessageHelper.Show(sb.ToString());
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Get deployment detail failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void GetWorkbookCompatibility(IWorkbookLogger logger)
        {
            try
            {
                var versionStatus = GetVersionStatus();
                MessageHelper.Show(versionStatus);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Get workbook version / compatibility failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal void ValidatePackageOnRibbon(IWorkbookLogger logger)
        {
            try
            {
                var sb = ValidatePackage();
                if (sb.Length == 0)
                {
                    MessageHelper.Show($"No {BexConstants.DataValidationTitle.ToLower()} messages", MessageType.Success);
                }
                else
                {
                    MessageHelper.Show(BexConstants.DataValidationTitle, sb.ToString(), MessageType.Warning);
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"{BexConstants.DataValidationTitle.ToStartOfSentence()} routine failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public StringBuilder ValidatePackage()
        {
            var sb  = new StringBuilder();

            var packageModel = Package.CreatePackageModel();
            if (Package.ExcelValidation.Length > 0)
            {
                sb.Append(Package.ExcelValidation);
                return sb;
            }

            var validation = packageModel.Validate();
            if (validation.Length <= 0) return sb;

            sb.Append(validation);
            return sb;
        }

        public bool ValidatePackageWithMessage(string name)
        {
            var validationMessage = ValidatePackage();
            if (validationMessage.Length == 0) return true;

            var message = new StringBuilder();
            message.AppendLine(
                $"The {name.ToLower()} can't be run because {BexConstants.DataValidationTitle.ToLower()} failed.");
            message.AppendLine();
            message.Append(validationMessage);
            MessageHelper.Show(name, message.ToString(), MessageType.Warning);
            return false;
        }

        private void DeleteSegment(ISegment segment)
        {
            var displayOrder = segment.DisplayOrder;
            segment.WorksheetManager.DeleteWorksheet();
            Package.Segments.Remove(segment);

            SelectSegment(displayOrder - 1);

            var summaryBuilder = new ProspectiveExposureSummaryBuilder();
            summaryBuilder.Build();
        }

        private void SelectSegment(int displayOrder)
        {
            var segment = Package.GetSegmentBasedOnDisplayOrder(displayOrder);
            if (segment == null)
            {
                Package.IsSelected = true;
                return;
            }
            segment.IsSelected = true;
        }

        internal void PerformQualityControl(IWorkbookLogger logger)
        {
            var packageModel = Package.CreatePackageModelForQualityControl();

            try
            {
                var qualityControlMessage = packageModel.PerformQualityControl();
                if (qualityControlMessage.Length > 0)
                {
                    MessageHelper.Show(BexConstants.DataQualityControlTitle, qualityControlMessage.ToString(), MessageType.Warning);
                    return;
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var errorMessage = $"{BexConstants.DataQualityControlTitle.ToStartOfSentence()} failed"; 
                MessageHelper.Show(errorMessage, MessageType.Stop);
                return;
            }

            MessageHelper.Show(BexConstants.DataQualityControlTitle, $"No {BexConstants.DataQualityControlTitle.ToLower()} messages", MessageType.Success);
        }
    }
}