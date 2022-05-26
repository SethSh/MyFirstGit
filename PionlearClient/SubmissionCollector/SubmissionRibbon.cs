using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelWorkspaceFolder;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector
{
    public partial class SubmissionRibbon
    {
        private Task _waitForRibbonTask;

        private void SubmissionRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (_waitForRibbonTask == null || _waitForRibbonTask.IsCompleted)
            {
                _waitForRibbonTask = SelectSubmissionRibbonWithWait();
            }
        }

        public Task SelectSubmissionRibbonWithWait()
        {
            //just enough time to allow workbook activate to end and redraw submission ribbon
            var task = new Task(() =>
            {
                const int threadDelay = 500;
                const long failTime = 10000;
                bool success;

                long delayTime = 0;
                // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                do
                {
                    success = SelectSubmissionRibbon();
                    if (success) return;
                    //ToDo investigate throw when closing bex wb and non-Bex wb is then activated
                    Thread.Sleep(threadDelay);
                    delayTime += threadDelay;
                } while (!success || delayTime <= failTime);
            });
            task.Start();
            return task;
        }

        public bool SelectSubmissionRibbon()
        {
            try
            {
                var ribbonTab = SubmissionTab;
                // ReSharper disable once RedundantCast
                var ribbonUi = (IRibbonUI) ribbonTab.RibbonUI;
                ribbonUi?.ActivateTab("SubmissionTab");
                return true;
            }
            catch (ArgumentException ex)
            {
                Debug.WriteLine(ex.Message);
                return false;
            }
        }

        private void DuplicateRiskClassButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.ThisExcelWorkspace.DuplicateSelectedSegment(new StackTraceLogger());
        }

        private void TransposeRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new RangeTransposeManager();
            manager.Transpose(new StackTraceLogger());
        }

        private void CreateJsonButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            SerializeManager.ConvertModelToJson(Globals.ThisWorkbook.ThisExcelWorkspace.Package, new StackTraceLogger());
        }

        private void LoadSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            UploadButton_Click(sender, e);
        }

        private void UploadButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

                var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
                var uploaderManager = new WorkbookUploaderManager(new WorkbookUploader(package));
                uploaderManager.Manage(package);
            }
            catch (Exception ex)
            {
                var logger = new StackTraceLogger();
                logger.WriteNew(ex);
                var message = $"{BexConstants.UploadName.ToStartOfSentence()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
            finally
            {
                Globals.ThisWorkbook.IsCurrentlyUploading = false;
            }
        }

        private void DeletePackageButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new ServerPackageDeleteManager();
            manager.DeleteWrapper(Globals.ThisWorkbook.ThisExcelWorkspace.Package, new StackTraceLogger());
        }


        private void ShowCustomXmlPartButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            CustomXmlPartManager.GetBexCustomXmlPartWithMessage(new StackTraceLogger());
        }

        private void InsertRowsIntoRangeSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            InsertRowsIntoRangeButton_Click(sender, e);
        }

        private void InsertRowsIntoRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.InsertRowsIntoRange(new StackTraceLogger());
        }

        private void InsertRowsIntoRangeAlternativeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.InsertRowsIntoRangeWithCount(new StackTraceLogger());
        }

        private void DeleteRowsFromRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.DeleteRowsFromRange(new StackTraceLogger());
        }

        private void InsertColumnsIntoRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.InsertColumnsIntoRange(new StackTraceLogger());
        }

        private void DeleteColumnsFromRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.DeleteColumnsFromRange(new StackTraceLogger());
        }

        private void DecouplePackageButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.ThisExcelWorkspace.Package.DecouplePackage(new StackTraceLogger());
        }

        private void RefreshReferenceDataButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            ReferenceDataManager.ReloadWithProgressBar(new StackTraceLogger());
        }

        private void ShowServerPackageAsJsonButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            SerializeManager.GetPackageFromBex(new StackTraceLogger());
        }

        private void AddWorksheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            CustomWorksheetManager.AddWorksheet(new StackTraceLogger());
        }

        private void DeleteWorksheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            CustomWorksheetManager.DeleteWorksheet(new StackTraceLogger());
        }

        private void SortByStateNameButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new StateSortManager();
            manager.SortByStateName(new StackTraceLogger());
        }

        private void SortByStateIdButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new StateSortManager();
            manager.SortByStateCode(new StackTraceLogger());
        }

        private void SortByStateIdWithCwOnTopButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new StateSortManager();
            manager.SortByStateCodeWithCwOnTop(new StackTraceLogger());
        }

        private void SortByStateNameWithCwOnTopSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new StateSortManager();
            manager.SortByStateNameWithCwOnTop(new StackTraceLogger());
        }

        private void DecouplePackageSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            DecouplePackageButton_Click(sender, e);
        }

        private void GetUrlsButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
                StringBuilder sb;
                using (new CursorToWait())
                {
                    sb = new StringBuilder();
                    sb.AppendLine("Bex");
                    sb.AppendLine(ConfigurationHelper.BexBaseUrl);
                    sb.AppendLine();
                    sb.AppendLine("UWPF Token");
                    sb.AppendLine(ConfigurationHelper.UwpfTokenUrl);
                    sb.AppendLine();
                    sb.AppendLine("Global Key Data");
                    sb.AppendLine(ConfigurationHelper.KeyDataBaseUrl);
                    sb.AppendLine();
                    sb.AppendLine("Submission API");
                    sb.AppendLine(ConfigurationHelper.BexSubmissionsUrl);
                }

                MessageHelper.Show("Services URLs", sb.ToString());
            }
            catch (Exception ex)
            {
                var logger = new StackTraceLogger();
                logger.WriteNew(ex);
                const string message = "Get URLs failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private void RenameWorksheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            CustomWorksheetManager.RenameWorksheet(new StackTraceLogger());
        }

        private void MoveWorksheetLeftButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            CustomWorksheetManager.MoveWorksheetLeft(new StackTraceLogger());
        }

        private void MoveWorksheetRightButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            CustomWorksheetManager.MoveWorksheetRight(new StackTraceLogger());
        }

        private void SetPackageIsDirtyTrueButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var setter = new DirtyFlagSetter();
            setter.SetPackageIsDirty(Globals.ThisWorkbook.ThisExcelWorkspace.Package, true);

            var synchronizationManager = new SynchronizationManager();
            synchronizationManager.ShowIndicators();
        }

        private void SetPackageIsDirtyFalseButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var dirtyFlagSetter = new DirtyFlagSetter();
            dirtyFlagSetter.SetPackageIsDirty(Globals.ThisWorkbook.ThisExcelWorkspace.Package, false);

            var synchronizationManager = new SynchronizationManager();
            synchronizationManager.ShowIndicators();
        }


        private void PaneToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.TogglePane(
                // ReSharper disable once RedundantCast
                (Microsoft.Office.Interop.Excel.Application) Globals.ThisWorkbook.Application,
                PaneToggleButton,
                new StackTraceLogger());

        }

        private void AddSegmentSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            AddSegmentButton_Click(sender, e);
        }

        private void AddSegmentButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            Globals.ThisWorkbook.ThisExcelWorkspace.AddSegment(new StackTraceLogger());
        }

        private void DeleteSegmentButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            Globals.ThisWorkbook.ThisExcelWorkspace.DeleteSegment(new StackTraceLogger());
        }

        private void PolicyProfileDimensionToFlatButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new PolicyProfileDimensionManager();
            manager.ChangeToFlat(new StackTraceLogger());
        }

        private void PolicyProfileDimensionToLimitBySirButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new PolicyProfileDimensionManager();
            manager.ChangeToLimitBySir(new StackTraceLogger());
        }

        private void PolicyProfileDimensionToSirByLimitButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new PolicyProfileDimensionManager();
            manager.ChangeToSirByLimit(new StackTraceLogger());
        }

        private void ShowPolicyProfileDimensionHelp_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new PolicyProfileDimensionHelpManager();
            manager.Help(new StackTraceLogger());
        }

        private void QuickFillDatesWithValuesSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            QuickFillDatesWithValuesButton_Click(sender, e);
        }

        private void QuickFillDatesWithValuesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = FillManagerFactory.Create(new StackTraceLogger());
            manager?.Fill(new StackTraceLogger(), true);
        }

        private void QuickFillDatesWithFormulasButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = FillManagerFactory.Create(new StackTraceLogger());
            manager?.Fill(new StackTraceLogger(), false);
        }

        private void SublineWizardButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var sublineManager = new SublineWizardManager();
            sublineManager.ModifySublinesWithWizard(new StackTraceLogger());
        }

        private void ShowIsDirtyFlagsButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var synchronizationManager = new SynchronizationManager();
            synchronizationManager.ShowIndicators();
        }

        private void ShowDirtyFlagsSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            ShowIsDirtyFlagsButton_Click(sender, e);
        }

        private void UserPreferencesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var modifier = new UserPreferenceModifier();
            modifier.Modify(new StackTraceLogger());
        }

        private void DecoupleComponentButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new DecoupleManager();
            manager.DecoupleComponent(new StackTraceLogger());
        }

        private void DecoupleSegmentButtonClick(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var manager = new DecoupleManager();
            manager.DecoupleSegmentSelected(new StackTraceLogger());
        }

        private void GetStartedButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            GetStartedDocumentationManager.GetDocumentation(new StackTraceLogger());
        }

        private void UserPreferencesSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            UserPreferencesButton_Click(sender, e);
        }

        private void UserPreferencesResetButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            UserPreferencesManager.ResetUserPreferences(new StackTraceLogger());
        }

        private void UmbrellaPolicyProfileButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            UmbrellaTypePolicyProfileWizardManager.ModifyUmbrellaType(new StackTraceLogger());
        }

        private void MoveRangeRightButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            RangeMover.MoveRight(new StackTraceLogger());
        }

        private void MoveRangeLeftButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            RangeMover.MoveLeft(new StackTraceLogger());
        }

        private void BuildVersionButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.ThisExcelWorkspace.GetBuildVersion(new StackTraceLogger());
        }

        private void UploadPreviewButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            if (!Globals.ThisWorkbook.ThisExcelWorkspace.ValidatePackageWithMessage(BexConstants.UploadPreviewName)) return;

            var synchronizationRenderer = new SynchronizationManager();
            synchronizationRenderer.RenderInBackground(new StackTraceLogger());
        }

        private void ShowUserPreferencesAsJsonButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            UserPreferencesManager.ShowJson(new StackTraceLogger());
        }

        private void ValidateButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.ThisExcelWorkspace.ValidatePackageOnRibbon(new StackTraceLogger());
        }

        private void LedgerSourceIdButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var renderer = new LedgerRenderer();
            renderer.ShowIsDirtyFlags(Globals.ThisWorkbook.ThisExcelWorkspace.Package, new StackTraceLogger());
        }

        private void RenewSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            RenewButton_Click(sender, e);
        }

        private void RenewButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            RenewManager.Renew(new StackTraceLogger());
        }
        
        private void RenewHelpButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            RenewDocumentationManager.Document(new StackTraceLogger());
        }

        private void QualityControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            if (!Globals.ThisWorkbook.ThisExcelWorkspace.ValidatePackageWithMessage(BexConstants.DataQualityControlTitle)) return;

            Globals.ThisWorkbook.ThisExcelWorkspace.PerformQualityControl(new StackTraceLogger());
        }

        private void DeleteRenewProperties_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            package.DeleteRenewPropertiesAndPredecessorIds(new StackTraceLogger());

        }

        private void QualityControlSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            QualityControlButton_Click(sender, e);
        }

        private void QualityControlDocumentationButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            QualityControlDocumentationManager.Document(new StackTraceLogger());
        }

        private void IsCompatibleButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            Globals.ThisWorkbook.ThisExcelWorkspace.GetWorkbookCompatibility(new StackTraceLogger());
        }

        private void ServerHistoryButton2_Click(object sender, RibbonControlEventArgs e)
        {
            ServerHistoryManager.Show(Globals.ThisWorkbook.ThisExcelWorkspace.Package, new StackTraceLogger());
        }

        private void ShowRatingAnalysisButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            RatingAnalysisManager.ShowRatingAnalyses(Globals.ThisWorkbook.ThisExcelWorkspace.Package, new StackTraceLogger());
        }

        private void FactorsButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            FactorsWorksheetCreator.Create(new StackTraceLogger());
        }

        private void RevertFormatButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            Globals.ThisWorkbook.ReformatDataComponent(new StackTraceLogger());
        }

        private void EstimateFormatButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            Globals.ThisWorkbook.ThisExcelWorkspace.ToggleEstimate(new StackTraceLogger());
        }

        private void UmbrellaSelectorButton_Click(object sender, RibbonControlEventArgs e)
        {
            var umbrellaSelector = new UmbrellaTypeAllocatorManager();
            umbrellaSelector.SelectUmbrellaTypes(new StackTraceLogger());
        }

        private void UmbrellaSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            UmbrellaSelectorButton_Click(sender, e);
        }

        private void AppDataButton_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start(ConfigurationHelper.AppDataFolder);
        }

        private void ShowRangeNamesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            RangeNameDisplayer.SetVisibility(true, new StackTraceLogger());
        }

        private void HideRangeNamesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            RangeNameDisplayer.SetVisibility(false, new StackTraceLogger());
        }

        private void ReplaceButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            FindAndReplacer.FindAndReplace(new StackTraceLogger());
        }

        private void FixCheckboxSynchronizationButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            CheckboxSynchronizationFixer.Fix(new StackTraceLogger());
        }

        private void RebuildButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

                var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
                var bexCommunicationManager = new BexCommunicationManager();
                var rebuilder = new WorkbookRebuilder(bexCommunicationManager, new StackTraceLogger());

                var rebuilderManager = new WorkbookRebuilderManager(rebuilder);
                rebuilderManager.Manage(package);
            }
            catch (Exception ex)
            {
                var logger = new StackTraceLogger();
                logger.WriteNew(ex);
                var message = $"{BexConstants.RebuildName.ToStartOfSentence()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private void ChangeTivRange_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;

            TotalInsuredValueExpandManager.ChangeExpand(new StackTraceLogger());
        }

        private void ChangeWorkersCompClassCodeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            WorkersCompClassCodeActivator.ToggleUtilization(new StackTraceLogger());
        }

        private void WorkerCompClassCodeMappingButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            WorkersCompClassCodeMapper.Map(new StackTraceLogger());
        }

        private void ShowWorkersCompClassCodesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            WorkersCompClassCodeReferenceManager.Render(new StackTraceLogger());
        }

        private void ShowWorkersCompClassCodesQueryButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            WorkersCompClassCodeReferenceSearchManager.Render(new StackTraceLogger());
        }

        private void ChangeWorkersCompStateByHazardGroupButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisWorkbook.AssertNotEditingCells()) return;
            WorkersCompStateByHazardGroupIndependentActivator.ToggleIndependence(new StackTraceLogger());
        }
    }
}
