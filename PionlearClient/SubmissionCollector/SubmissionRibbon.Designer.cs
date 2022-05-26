using SubmissionCollector.Properties;

namespace SubmissionCollector
{
    partial class SubmissionRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SubmissionRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        public void SetSynchronizationButtonImage(bool isDirty)
        {
            if (isDirty)
            {
                this.UploadSplitButton.Image = Resources.CloudRed;
                this.UploadButton.Image = Resources.CloudRed;
            }
            else
            {
                this.UploadSplitButton.Image = Resources.CloudChecked;
                this.UploadButton.Image = Resources.CloudChecked;
            }
            Globals.ThisWorkbook.LoadButtonDirtyState = isDirty;
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.SubmissionTab = this.Factory.CreateRibbonTab();
            this.InventoryGroup = this.Factory.CreateRibbonGroup();
            this.SegmentGroup = this.Factory.CreateRibbonGroup();
            this.RangeUtilitiesGroup = this.Factory.CreateRibbonGroup();
            this.UtilitiesGroup = this.Factory.CreateRibbonGroup();
            this.DatabaseGroup = this.Factory.CreateRibbonGroup();
            this.WorkbookGroup = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.AdminGroup = this.Factory.CreateRibbonGroup();
            this.UserPreferencesButtonGroup = this.Factory.CreateRibbonGroup();
            this.PaneToggleButton = this.Factory.CreateRibbonToggleButton();
            this.AddSegmentSplitButton = this.Factory.CreateRibbonSplitButton();
            this.AddSegmentButton = this.Factory.CreateRibbonButton();
            this.DeleteSegmentButton = this.Factory.CreateRibbonButton();
            this.DuplicateRiskClassButton = this.Factory.CreateRibbonButton();
            this.SublineWizardButton = this.Factory.CreateRibbonButton();
            this.UmbrellaSplitButton = this.Factory.CreateRibbonSplitButton();
            this.UmbrellaSelectorButton = this.Factory.CreateRibbonButton();
            this.UmbrellaPolicyProfileButton = this.Factory.CreateRibbonButton();
            this.ComponentArrangeMenu = this.Factory.CreateRibbonMenu();
            this.MoveComponentsMenu = this.Factory.CreateRibbonMenu();
            this.MoveRangeRightButton = this.Factory.CreateRibbonButton();
            this.MoveRangeLeftButton = this.Factory.CreateRibbonButton();
            this.TransposeMatrixButton = this.Factory.CreateRibbonButton();
            this.PolicyProfileDimensionMenu = this.Factory.CreateRibbonMenu();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.ChangeTivRange = this.Factory.CreateRibbonButton();
            this.FormatSplitButton = this.Factory.CreateRibbonSplitButton();
            this.ReformatterButton = this.Factory.CreateRibbonButton();
            this.EstimateButton = this.Factory.CreateRibbonButton();
            this.InsertRowsIntoRangeSplitButton = this.Factory.CreateRibbonSplitButton();
            this.InsertRowsIntoRangeButton = this.Factory.CreateRibbonButton();
            this.InsertRowsIntoRangeAlternativeButton = this.Factory.CreateRibbonButton();
            this.DeleteRowsFromRangeButton = this.Factory.CreateRibbonButton();
            this.InsertColumnsIntoRangeButton = this.Factory.CreateRibbonButton();
            this.DeleteColumnsFromRangeButton = this.Factory.CreateRibbonButton();
            this.StateSorterMenu = this.Factory.CreateRibbonMenu();
            this.SortByStateNameWithCwOnTopButton = this.Factory.CreateRibbonButton();
            this.SortByStateNameButton = this.Factory.CreateRibbonButton();
            this.SortByStateIdWithCwOnTopButton = this.Factory.CreateRibbonButton();
            this.SortByStateIdButton = this.Factory.CreateRibbonButton();
            this.QuickFillDatesWithValuesSplitButton = this.Factory.CreateRibbonSplitButton();
            this.QuickFillDatesWithValuesButton = this.Factory.CreateRibbonButton();
            this.QuickFillDatesWithFormulasButton = this.Factory.CreateRibbonButton();
            this.ReplaceButton = this.Factory.CreateRibbonButton();
            this.CustomWorksheetMenu = this.Factory.CreateRibbonMenu();
            this.AddWorksheetButton = this.Factory.CreateRibbonButton();
            this.DeleteWorksheetButton = this.Factory.CreateRibbonButton();
            this.RenameWorksheetButton = this.Factory.CreateRibbonButton();
            this.MoveWorksheetLeftButtton = this.Factory.CreateRibbonButton();
            this.MoveWorksheetRightButtton = this.Factory.CreateRibbonButton();
            this.FactorsButton = this.Factory.CreateRibbonButton();
            this.UploadSplitButton = this.Factory.CreateRibbonSplitButton();
            this.UploadButton = this.Factory.CreateRibbonButton();
            this.UploadPreviewButton = this.Factory.CreateRibbonButton();
            this.DeleteFromDatabaseButton = this.Factory.CreateRibbonButton();
            this.ShowRatingAnalysesButton = this.Factory.CreateRibbonButton();
            this.ServerHistoryButton = this.Factory.CreateRibbonButton();
            this.DecouplePackageSplitButton = this.Factory.CreateRibbonSplitButton();
            this.DecouplePackageButton = this.Factory.CreateRibbonButton();
            this.DecoupleSegmentButton = this.Factory.CreateRibbonButton();
            this.DecoupleComponentButton = this.Factory.CreateRibbonButton();
            this.ValidateButton = this.Factory.CreateRibbonButton();
            this.QualityControlSplitButton = this.Factory.CreateRibbonSplitButton();
            this.QualityControlButton = this.Factory.CreateRibbonButton();
            this.QualityControlDocumentationButton = this.Factory.CreateRibbonButton();
            this.RenewSplitButton = this.Factory.CreateRibbonSplitButton();
            this.RenewButton = this.Factory.CreateRibbonButton();
            this.DeleteRenewProperties = this.Factory.CreateRibbonButton();
            this.RenewHelpButton = this.Factory.CreateRibbonButton();
            this.WorkercCompMenu = this.Factory.CreateRibbonMenu();
            this.WorkersCompConfigureMenu = this.Factory.CreateRibbonMenu();
            this.ChangeWorkersCompClassCodeButton = this.Factory.CreateRibbonButton();
            this.ChangeWorkersCompStateByHazardGroupButton = this.Factory.CreateRibbonButton();
            this.WorkerCompClassCodeMappingButton = this.Factory.CreateRibbonButton();
            this.WorkersCompLookUpMenu = this.Factory.CreateRibbonMenu();
            this.ShowWorkersCompClassCodesButton = this.Factory.CreateRibbonButton();
            this.ShowWorkersCompClassCodesQueryButton = this.Factory.CreateRibbonButton();
            this.DevelopmentToolsMenu = this.Factory.CreateRibbonMenu();
            this.ShowServerPackageAsJsonButton = this.Factory.CreateRibbonButton();
            this.CreateJsonButton = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.ShowUserPrefsAsJsonButton = this.Factory.CreateRibbonButton();
            this.ShowIsDirtySplitButton = this.Factory.CreateRibbonSplitButton();
            this.ShoeIsDirtyButton = this.Factory.CreateRibbonButton();
            this.SetPackageIsDirtyTrueButton = this.Factory.CreateRibbonButton();
            this.SetPackageIsDirtyFalseButton = this.Factory.CreateRibbonButton();
            this.LedgerButton = this.Factory.CreateRibbonButton();
            this.ShowRangeNamesButton = this.Factory.CreateRibbonButton();
            this.HideRangeNamesButton = this.Factory.CreateRibbonButton();
            this.FixCheckboxSynchronizationButton = this.Factory.CreateRibbonButton();
            this.AppDataButton = this.Factory.CreateRibbonButton();
            this.RebuildWorkbookButton = this.Factory.CreateRibbonButton();
            this.UserPreferencesSplitButton = this.Factory.CreateRibbonSplitButton();
            this.UserPreferencesButton = this.Factory.CreateRibbonButton();
            this.UserPreferencesResetButton = this.Factory.CreateRibbonButton();
            this.RefreshReferenceDataButton = this.Factory.CreateRibbonButton();
            this.AboutMenu = this.Factory.CreateRibbonMenu();
            this.GetStartedButton = this.Factory.CreateRibbonButton();
            this.BuildVersionButton = this.Factory.CreateRibbonButton();
            this.IsCompatibleButton = this.Factory.CreateRibbonButton();
            this.GetUrlsButton = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.SubmissionTab.SuspendLayout();
            this.InventoryGroup.SuspendLayout();
            this.SegmentGroup.SuspendLayout();
            this.RangeUtilitiesGroup.SuspendLayout();
            this.UtilitiesGroup.SuspendLayout();
            this.DatabaseGroup.SuspendLayout();
            this.WorkbookGroup.SuspendLayout();
            this.AdminGroup.SuspendLayout();
            this.UserPreferencesButtonGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // SubmissionTab
            // 
            this.SubmissionTab.Groups.Add(this.InventoryGroup);
            this.SubmissionTab.Groups.Add(this.SegmentGroup);
            this.SubmissionTab.Groups.Add(this.RangeUtilitiesGroup);
            this.SubmissionTab.Groups.Add(this.UtilitiesGroup);
            this.SubmissionTab.Groups.Add(this.DatabaseGroup);
            this.SubmissionTab.Groups.Add(this.WorkbookGroup);
            this.SubmissionTab.Groups.Add(this.AdminGroup);
            this.SubmissionTab.Groups.Add(this.UserPreferencesButtonGroup);
            this.SubmissionTab.KeyTip = "S";
            this.SubmissionTab.Label = "SUBMISSION";
            this.SubmissionTab.Name = "SubmissionTab";
            // 
            // InventoryGroup
            // 
            this.InventoryGroup.Items.Add(this.PaneToggleButton);
            this.InventoryGroup.Name = "InventoryGroup";
            // 
            // SegmentGroup
            // 
            this.SegmentGroup.Items.Add(this.AddSegmentSplitButton);
            this.SegmentGroup.Items.Add(this.SublineWizardButton);
            this.SegmentGroup.Items.Add(this.UmbrellaSplitButton);
            this.SegmentGroup.Label = "Submission Segment";
            this.SegmentGroup.Name = "SegmentGroup";
            // 
            // RangeUtilitiesGroup
            // 
            this.RangeUtilitiesGroup.Items.Add(this.ComponentArrangeMenu);
            this.RangeUtilitiesGroup.Items.Add(this.FormatSplitButton);
            this.RangeUtilitiesGroup.Items.Add(this.WorkercCompMenu);
            this.RangeUtilitiesGroup.Items.Add(this.InsertRowsIntoRangeSplitButton);
            this.RangeUtilitiesGroup.Items.Add(this.StateSorterMenu);
            this.RangeUtilitiesGroup.Items.Add(this.QuickFillDatesWithValuesSplitButton);
            this.RangeUtilitiesGroup.Items.Add(this.ReplaceButton);
            this.RangeUtilitiesGroup.Label = "Range Utilities";
            this.RangeUtilitiesGroup.Name = "RangeUtilitiesGroup";
            // 
            // UtilitiesGroup
            // 
            this.UtilitiesGroup.Items.Add(this.CustomWorksheetMenu);
            this.UtilitiesGroup.Label = "Workbook Utilities";
            this.UtilitiesGroup.Name = "UtilitiesGroup";
            // 
            // DatabaseGroup
            // 
            this.DatabaseGroup.Items.Add(this.UploadSplitButton);
            this.DatabaseGroup.Items.Add(this.DeleteFromDatabaseButton);
            this.DatabaseGroup.Items.Add(this.ShowRatingAnalysesButton);
            this.DatabaseGroup.Items.Add(this.ServerHistoryButton);
            this.DatabaseGroup.Items.Add(this.DecouplePackageSplitButton);
            this.DatabaseGroup.Label = "Database";
            this.DatabaseGroup.Name = "DatabaseGroup";
            // 
            // WorkbookGroup
            // 
            this.WorkbookGroup.Items.Add(this.ValidateButton);
            this.WorkbookGroup.Items.Add(this.QualityControlSplitButton);
            this.WorkbookGroup.Items.Add(this.separator1);
            this.WorkbookGroup.Items.Add(this.RenewSplitButton);
            this.WorkbookGroup.Label = "Workbook";
            this.WorkbookGroup.Name = "WorkbookGroup";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // AdminGroup
            // 
            this.AdminGroup.Items.Add(this.DevelopmentToolsMenu);
            this.AdminGroup.Items.Add(this.AppDataButton);
            this.AdminGroup.Items.Add(this.RebuildWorkbookButton);
            this.AdminGroup.Label = "Administrative";
            this.AdminGroup.Name = "AdminGroup";
            // 
            // UserPreferencesButtonGroup
            // 
            this.UserPreferencesButtonGroup.Items.Add(this.UserPreferencesSplitButton);
            this.UserPreferencesButtonGroup.Items.Add(this.AboutMenu);
            this.UserPreferencesButtonGroup.Label = "Add-In Management";
            this.UserPreferencesButtonGroup.Name = "UserPreferencesButtonGroup";
            // 
            // PaneToggleButton
            // 
            this.PaneToggleButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PaneToggleButton.Image = global::SubmissionCollector.Properties.Resources.Pane;
            this.PaneToggleButton.KeyTip = "I";
            this.PaneToggleButton.Label = "Inventory Pane";
            this.PaneToggleButton.Name = "PaneToggleButton";
            this.PaneToggleButton.ShowImage = true;
            this.PaneToggleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PaneToggleButton_Click);
            // 
            // AddSegmentSplitButton
            // 
            this.AddSegmentSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddSegmentSplitButton.Image = global::SubmissionCollector.Properties.Resources.Add;
            this.AddSegmentSplitButton.Items.Add(this.AddSegmentButton);
            this.AddSegmentSplitButton.Items.Add(this.DeleteSegmentButton);
            this.AddSegmentSplitButton.Items.Add(this.DuplicateRiskClassButton);
            this.AddSegmentSplitButton.KeyTip = "S";
            this.AddSegmentSplitButton.Label = "Add";
            this.AddSegmentSplitButton.Name = "AddSegmentSplitButton";
            this.AddSegmentSplitButton.ScreenTip = "Add Submission Segment";
            this.AddSegmentSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddSegmentSplitButton_Click);
            // 
            // AddSegmentButton
            // 
            this.AddSegmentButton.Image = global::SubmissionCollector.Properties.Resources.Add;
            this.AddSegmentButton.Label = "Add";
            this.AddSegmentButton.Name = "AddSegmentButton";
            this.AddSegmentButton.ScreenTip = "Add Submission Segment";
            this.AddSegmentButton.ShowImage = true;
            this.AddSegmentButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddSegmentButton_Click);
            // 
            // DeleteSegmentButton
            // 
            this.DeleteSegmentButton.Image = global::SubmissionCollector.Properties.Resources.DeleteX;
            this.DeleteSegmentButton.Label = "Delete";
            this.DeleteSegmentButton.Name = "DeleteSegmentButton";
            this.DeleteSegmentButton.ScreenTip = "Delete Submission Segment";
            this.DeleteSegmentButton.ShowImage = true;
            this.DeleteSegmentButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteSegmentButton_Click);
            // 
            // DuplicateRiskClassButton
            // 
            this.DuplicateRiskClassButton.Image = global::SubmissionCollector.Properties.Resources.Copy;
            this.DuplicateRiskClassButton.KeyTip = "D";
            this.DuplicateRiskClassButton.Label = "Duplicate";
            this.DuplicateRiskClassButton.Name = "DuplicateRiskClassButton";
            this.DuplicateRiskClassButton.ScreenTip = "Duplicate Submission Segment";
            this.DuplicateRiskClassButton.ShowImage = true;
            this.DuplicateRiskClassButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DuplicateRiskClassButton_Click);
            // 
            // SublineWizardButton
            // 
            this.SublineWizardButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SublineWizardButton.Image = global::SubmissionCollector.Properties.Resources.Wand;
            this.SublineWizardButton.KeyTip = "B";
            this.SublineWizardButton.Label = "Sublines Wizard";
            this.SublineWizardButton.Name = "SublineWizardButton";
            this.SublineWizardButton.ShowImage = true;
            this.SublineWizardButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SublineWizardButton_Click);
            // 
            // UmbrellaSplitButton
            // 
            this.UmbrellaSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UmbrellaSplitButton.Image = global::SubmissionCollector.Properties.Resources.Umbrella;
            this.UmbrellaSplitButton.Items.Add(this.UmbrellaSelectorButton);
            this.UmbrellaSplitButton.Items.Add(this.UmbrellaPolicyProfileButton);
            this.UmbrellaSplitButton.KeyTip = "L";
            this.UmbrellaSplitButton.Label = "Umbrella Type Selector";
            this.UmbrellaSplitButton.Name = "UmbrellaSplitButton";
            this.UmbrellaSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UmbrellaSplitButton_Click);
            // 
            // UmbrellaSelectorButton
            // 
            this.UmbrellaSelectorButton.Image = global::SubmissionCollector.Properties.Resources.Umbrella;
            this.UmbrellaSelectorButton.Label = "Selector";
            this.UmbrellaSelectorButton.Name = "UmbrellaSelectorButton";
            this.UmbrellaSelectorButton.ShowImage = true;
            this.UmbrellaSelectorButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UmbrellaSelectorButton_Click);
            // 
            // UmbrellaPolicyProfileButton
            // 
            this.UmbrellaPolicyProfileButton.Image = global::SubmissionCollector.Properties.Resources.Wand;
            this.UmbrellaPolicyProfileButton.Label = "Policy Profile Wizard";
            this.UmbrellaPolicyProfileButton.Name = "UmbrellaPolicyProfileButton";
            this.UmbrellaPolicyProfileButton.ShowImage = true;
            this.UmbrellaPolicyProfileButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UmbrellaPolicyProfileButton_Click);
            // 
            // ComponentArrangeMenu
            // 
            this.ComponentArrangeMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ComponentArrangeMenu.Image = global::SubmissionCollector.Properties.Resources.Arrange;
            this.ComponentArrangeMenu.Items.Add(this.MoveComponentsMenu);
            this.ComponentArrangeMenu.Items.Add(this.TransposeMatrixButton);
            this.ComponentArrangeMenu.Items.Add(this.PolicyProfileDimensionMenu);
            this.ComponentArrangeMenu.Items.Add(this.ChangeTivRange);
            this.ComponentArrangeMenu.KeyTip = "R";
            this.ComponentArrangeMenu.Label = "Reconfigure";
            this.ComponentArrangeMenu.Name = "ComponentArrangeMenu";
            this.ComponentArrangeMenu.ShowImage = true;
            // 
            // MoveComponentsMenu
            // 
            this.MoveComponentsMenu.Items.Add(this.MoveRangeRightButton);
            this.MoveComponentsMenu.Items.Add(this.MoveRangeLeftButton);
            this.MoveComponentsMenu.Label = "Move Range";
            this.MoveComponentsMenu.Name = "MoveComponentsMenu";
            this.MoveComponentsMenu.ShowImage = true;
            // 
            // MoveRangeRightButton
            // 
            this.MoveRangeRightButton.Image = global::SubmissionCollector.Properties.Resources.RightArrow;
            this.MoveRangeRightButton.KeyTip = "UR";
            this.MoveRangeRightButton.Label = "Move Right";
            this.MoveRangeRightButton.Name = "MoveRangeRightButton";
            this.MoveRangeRightButton.ShowImage = true;
            this.MoveRangeRightButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveRangeRightButton_Click);
            // 
            // MoveRangeLeftButton
            // 
            this.MoveRangeLeftButton.Image = global::SubmissionCollector.Properties.Resources.LeftArrow;
            this.MoveRangeLeftButton.KeyTip = "UL";
            this.MoveRangeLeftButton.Label = "Move Left";
            this.MoveRangeLeftButton.Name = "MoveRangeLeftButton";
            this.MoveRangeLeftButton.ShowImage = true;
            this.MoveRangeLeftButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveRangeLeftButton_Click);
            // 
            // TransposeMatrixButton
            // 
            this.TransposeMatrixButton.Label = "Transpose Range";
            this.TransposeMatrixButton.Name = "TransposeMatrixButton";
            this.TransposeMatrixButton.OfficeImageId = "AccessFormPivotTable";
            this.TransposeMatrixButton.ShowImage = true;
            this.TransposeMatrixButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TransposeRangeButton_Click);
            // 
            // PolicyProfileDimensionMenu
            // 
            this.PolicyProfileDimensionMenu.Image = global::SubmissionCollector.Properties.Resources.Change;
            this.PolicyProfileDimensionMenu.Items.Add(this.button6);
            this.PolicyProfileDimensionMenu.Items.Add(this.button7);
            this.PolicyProfileDimensionMenu.Items.Add(this.button8);
            this.PolicyProfileDimensionMenu.Items.Add(this.button9);
            this.PolicyProfileDimensionMenu.Label = "Liability: Toggle Policy Profile Dimension Options";
            this.PolicyProfileDimensionMenu.Name = "PolicyProfileDimensionMenu";
            this.PolicyProfileDimensionMenu.ShowImage = true;
            // 
            // button6
            // 
            this.button6.Label = "To Flat";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PolicyProfileDimensionToFlatButton_Click);
            // 
            // button7
            // 
            this.button7.Label = "To Limit by SIR";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PolicyProfileDimensionToLimitBySirButton_Click);
            // 
            // button8
            // 
            this.button8.Label = "To SIR by Limit";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PolicyProfileDimensionToSirByLimitButton_Click);
            // 
            // button9
            // 
            this.button9.Image = global::SubmissionCollector.Properties.Resources.Document;
            this.button9.Label = "Show Policy Profile Dimension Help";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowPolicyProfileDimensionHelp_Click);
            // 
            // ChangeTivRange
            // 
            this.ChangeTivRange.Label = "Property: Toggle TIV Range";
            this.ChangeTivRange.Name = "ChangeTivRange";
            this.ChangeTivRange.ScreenTip = "Include/exclude Share and Subject Policy Limit and Attachemnt";
            this.ChangeTivRange.ShowImage = true;
            this.ChangeTivRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeTivRange_Click);
            // 
            // FormatSplitButton
            // 
            this.FormatSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FormatSplitButton.Image = global::SubmissionCollector.Properties.Resources.PaintBrush;
            this.FormatSplitButton.Items.Add(this.ReformatterButton);
            this.FormatSplitButton.Items.Add(this.EstimateButton);
            this.FormatSplitButton.KeyTip = "F";
            this.FormatSplitButton.Label = "Format";
            this.FormatSplitButton.Name = "FormatSplitButton";
            this.FormatSplitButton.ScreenTip = "Revert format to original state";
            // 
            // ReformatterButton
            // 
            this.ReformatterButton.Image = global::SubmissionCollector.Properties.Resources.PaintBrush;
            this.ReformatterButton.Label = "Revert Format";
            this.ReformatterButton.Name = "ReformatterButton";
            this.ReformatterButton.ScreenTip = "Revert format to original state";
            this.ReformatterButton.ShowImage = true;
            this.ReformatterButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RevertFormatButton_Click);
            // 
            // EstimateButton
            // 
            this.EstimateButton.Image = global::SubmissionCollector.Properties.Resources.RedPen;
            this.EstimateButton.KeyTip = "UT";
            this.EstimateButton.Label = "Toggle Estimate";
            this.EstimateButton.Name = "EstimateButton";
            this.EstimateButton.OfficeImageId = "HappyFace";
            this.EstimateButton.ScreenTip = "Change range background color to identify estimate";
            this.EstimateButton.ShowImage = true;
            this.EstimateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EstimateFormatButton_Click);
            // 
            // InsertRowsIntoRangeSplitButton
            // 
            this.InsertRowsIntoRangeSplitButton.Checked = true;
            this.InsertRowsIntoRangeSplitButton.Image = global::SubmissionCollector.Properties.Resources.AddRow;
            this.InsertRowsIntoRangeSplitButton.Items.Add(this.InsertRowsIntoRangeButton);
            this.InsertRowsIntoRangeSplitButton.Items.Add(this.InsertRowsIntoRangeAlternativeButton);
            this.InsertRowsIntoRangeSplitButton.Items.Add(this.DeleteRowsFromRangeButton);
            this.InsertRowsIntoRangeSplitButton.Items.Add(this.InsertColumnsIntoRangeButton);
            this.InsertRowsIntoRangeSplitButton.Items.Add(this.DeleteColumnsFromRangeButton);
            this.InsertRowsIntoRangeSplitButton.KeyTip = "UI";
            this.InsertRowsIntoRangeSplitButton.Label = "Insert Rows";
            this.InsertRowsIntoRangeSplitButton.Name = "InsertRowsIntoRangeSplitButton";
            this.InsertRowsIntoRangeSplitButton.ScreenTip = "Insert rows within existing range";
            this.InsertRowsIntoRangeSplitButton.SuperTip = "Selection\'s top left cell determines which range";
            this.InsertRowsIntoRangeSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertRowsIntoRangeSplitButton_Click);
            // 
            // InsertRowsIntoRangeButton
            // 
            this.InsertRowsIntoRangeButton.Image = global::SubmissionCollector.Properties.Resources.AddRow;
            this.InsertRowsIntoRangeButton.Label = "Insert Rows";
            this.InsertRowsIntoRangeButton.Name = "InsertRowsIntoRangeButton";
            this.InsertRowsIntoRangeButton.ScreenTip = "Insert rows within existing range";
            this.InsertRowsIntoRangeButton.ShowImage = true;
            this.InsertRowsIntoRangeButton.SuperTip = "Selection\'s top left cell determines which range";
            this.InsertRowsIntoRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertRowsIntoRangeButton_Click);
            // 
            // InsertRowsIntoRangeAlternativeButton
            // 
            this.InsertRowsIntoRangeAlternativeButton.Image = global::SubmissionCollector.Properties.Resources.AddRow;
            this.InsertRowsIntoRangeAlternativeButton.Label = "Insert Rows With Count";
            this.InsertRowsIntoRangeAlternativeButton.Name = "InsertRowsIntoRangeAlternativeButton";
            this.InsertRowsIntoRangeAlternativeButton.ScreenTip = "Insert rows within existing range by specifying count";
            this.InsertRowsIntoRangeAlternativeButton.ShowImage = true;
            this.InsertRowsIntoRangeAlternativeButton.SuperTip = "Selection\'s top left cell determines which range.";
            this.InsertRowsIntoRangeAlternativeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertRowsIntoRangeAlternativeButton_Click);
            // 
            // DeleteRowsFromRangeButton
            // 
            this.DeleteRowsFromRangeButton.Image = global::SubmissionCollector.Properties.Resources.DeleteRow;
            this.DeleteRowsFromRangeButton.Label = "Delete Rows";
            this.DeleteRowsFromRangeButton.Name = "DeleteRowsFromRangeButton";
            this.DeleteRowsFromRangeButton.ScreenTip = "Delete rows from the existing range";
            this.DeleteRowsFromRangeButton.ShowImage = true;
            this.DeleteRowsFromRangeButton.SuperTip = "Selection\'s top left cell determines which range";
            this.DeleteRowsFromRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteRowsFromRangeButton_Click);
            // 
            // InsertColumnsIntoRangeButton
            // 
            this.InsertColumnsIntoRangeButton.Image = global::SubmissionCollector.Properties.Resources.AddColumn;
            this.InsertColumnsIntoRangeButton.Label = "Insert Columns";
            this.InsertColumnsIntoRangeButton.Name = "InsertColumnsIntoRangeButton";
            this.InsertColumnsIntoRangeButton.ScreenTip = "Expand range columns within the existing range";
            this.InsertColumnsIntoRangeButton.ShowImage = true;
            this.InsertColumnsIntoRangeButton.SuperTip = "Selection\'s top left cell determines which range";
            this.InsertColumnsIntoRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertColumnsIntoRangeButton_Click);
            // 
            // DeleteColumnsFromRangeButton
            // 
            this.DeleteColumnsFromRangeButton.Image = global::SubmissionCollector.Properties.Resources.DeleteColumn;
            this.DeleteColumnsFromRangeButton.Label = "Delete Columns";
            this.DeleteColumnsFromRangeButton.Name = "DeleteColumnsFromRangeButton";
            this.DeleteColumnsFromRangeButton.ScreenTip = "Delete columns from the exising range";
            this.DeleteColumnsFromRangeButton.ShowImage = true;
            this.DeleteColumnsFromRangeButton.SuperTip = "Selection\'s top left cell determines which range";
            this.DeleteColumnsFromRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteColumnsFromRangeButton_Click);
            // 
            // StateSorterMenu
            // 
            this.StateSorterMenu.Image = global::SubmissionCollector.Properties.Resources.SortAscending;
            this.StateSorterMenu.Items.Add(this.SortByStateNameWithCwOnTopButton);
            this.StateSorterMenu.Items.Add(this.SortByStateNameButton);
            this.StateSorterMenu.Items.Add(this.SortByStateIdWithCwOnTopButton);
            this.StateSorterMenu.Items.Add(this.SortByStateIdButton);
            this.StateSorterMenu.KeyTip = "US";
            this.StateSorterMenu.Label = "State Sorter";
            this.StateSorterMenu.Name = "StateSorterMenu";
            this.StateSorterMenu.ShowImage = true;
            // 
            // SortByStateNameWithCwOnTopButton
            // 
            this.SortByStateNameWithCwOnTopButton.Label = "Name (countrywide on top)";
            this.SortByStateNameWithCwOnTopButton.Name = "SortByStateNameWithCwOnTopButton";
            this.SortByStateNameWithCwOnTopButton.ShowImage = true;
            this.SortByStateNameWithCwOnTopButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SortByStateNameWithCwOnTopSplitButton_Click);
            // 
            // SortByStateNameButton
            // 
            this.SortByStateNameButton.Label = "Name";
            this.SortByStateNameButton.Name = "SortByStateNameButton";
            this.SortByStateNameButton.ShowImage = true;
            this.SortByStateNameButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SortByStateNameButton_Click);
            // 
            // SortByStateIdWithCwOnTopButton
            // 
            this.SortByStateIdWithCwOnTopButton.Label = "Abbreviation (CW on top)";
            this.SortByStateIdWithCwOnTopButton.Name = "SortByStateIdWithCwOnTopButton";
            this.SortByStateIdWithCwOnTopButton.ShowImage = true;
            this.SortByStateIdWithCwOnTopButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SortByStateIdWithCwOnTopButton_Click);
            // 
            // SortByStateIdButton
            // 
            this.SortByStateIdButton.Label = "Abbreviation";
            this.SortByStateIdButton.Name = "SortByStateIdButton";
            this.SortByStateIdButton.ShowImage = true;
            this.SortByStateIdButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SortByStateIdButton_Click);
            // 
            // QuickFillDatesWithValuesSplitButton
            // 
            this.QuickFillDatesWithValuesSplitButton.Image = global::SubmissionCollector.Properties.Resources.Calendar;
            this.QuickFillDatesWithValuesSplitButton.Items.Add(this.QuickFillDatesWithValuesButton);
            this.QuickFillDatesWithValuesSplitButton.Items.Add(this.QuickFillDatesWithFormulasButton);
            this.QuickFillDatesWithValuesSplitButton.KeyTip = "UQ";
            this.QuickFillDatesWithValuesSplitButton.Label = "Quick Fill Dates";
            this.QuickFillDatesWithValuesSplitButton.Name = "QuickFillDatesWithValuesSplitButton";
            this.QuickFillDatesWithValuesSplitButton.ScreenTip = "Writes date values into selected date range";
            this.QuickFillDatesWithValuesSplitButton.SuperTip = "Before quick filling, populate first row.";
            this.QuickFillDatesWithValuesSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QuickFillDatesWithValuesSplitButton_Click);
            // 
            // QuickFillDatesWithValuesButton
            // 
            this.QuickFillDatesWithValuesButton.Image = global::SubmissionCollector.Properties.Resources.Calendar;
            this.QuickFillDatesWithValuesButton.Label = "Quick Fill Dates";
            this.QuickFillDatesWithValuesButton.Name = "QuickFillDatesWithValuesButton";
            this.QuickFillDatesWithValuesButton.ScreenTip = "Writes date values into selected date range";
            this.QuickFillDatesWithValuesButton.ShowImage = true;
            this.QuickFillDatesWithValuesButton.SuperTip = "Before quick filling, populate first row.";
            this.QuickFillDatesWithValuesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QuickFillDatesWithValuesButton_Click);
            // 
            // QuickFillDatesWithFormulasButton
            // 
            this.QuickFillDatesWithFormulasButton.Image = global::SubmissionCollector.Properties.Resources.Calendar;
            this.QuickFillDatesWithFormulasButton.Label = "Quick Fill Dates With Formulas";
            this.QuickFillDatesWithFormulasButton.Name = "QuickFillDatesWithFormulasButton";
            this.QuickFillDatesWithFormulasButton.ScreenTip = "Writes date formulas into selected date range";
            this.QuickFillDatesWithFormulasButton.ShowImage = true;
            this.QuickFillDatesWithFormulasButton.SuperTip = "Before quick filling, populate first row.";
            this.QuickFillDatesWithFormulasButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QuickFillDatesWithFormulasButton_Click);
            // 
            // ReplaceButton
            // 
            this.ReplaceButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ReplaceButton.Image = global::SubmissionCollector.Properties.Resources.Replace;
            this.ReplaceButton.KeyTip = "UR";
            this.ReplaceButton.Label = "Find and Replace";
            this.ReplaceButton.Name = "ReplaceButton";
            this.ReplaceButton.ShowImage = true;
            this.ReplaceButton.Tag = "Find and replace on protected worksheets";
            this.ReplaceButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReplaceButton_Click);
            // 
            // CustomWorksheetMenu
            // 
            this.CustomWorksheetMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CustomWorksheetMenu.Image = global::SubmissionCollector.Properties.Resources.Spreadsheet;
            this.CustomWorksheetMenu.Items.Add(this.AddWorksheetButton);
            this.CustomWorksheetMenu.Items.Add(this.DeleteWorksheetButton);
            this.CustomWorksheetMenu.Items.Add(this.RenameWorksheetButton);
            this.CustomWorksheetMenu.Items.Add(this.MoveWorksheetLeftButtton);
            this.CustomWorksheetMenu.Items.Add(this.MoveWorksheetRightButtton);
            this.CustomWorksheetMenu.Items.Add(this.FactorsButton);
            this.CustomWorksheetMenu.KeyTip = "C";
            this.CustomWorksheetMenu.Label = "Custom Worksheet";
            this.CustomWorksheetMenu.Name = "CustomWorksheetMenu";
            this.CustomWorksheetMenu.ShowImage = true;
            // 
            // AddWorksheetButton
            // 
            this.AddWorksheetButton.Image = global::SubmissionCollector.Properties.Resources.Add;
            this.AddWorksheetButton.Label = "Add";
            this.AddWorksheetButton.Name = "AddWorksheetButton";
            this.AddWorksheetButton.ShowImage = true;
            this.AddWorksheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddWorksheetButton_Click);
            // 
            // DeleteWorksheetButton
            // 
            this.DeleteWorksheetButton.Image = global::SubmissionCollector.Properties.Resources.DeleteX;
            this.DeleteWorksheetButton.Label = "Delete";
            this.DeleteWorksheetButton.Name = "DeleteWorksheetButton";
            this.DeleteWorksheetButton.ShowImage = true;
            this.DeleteWorksheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteWorksheetButton_Click);
            // 
            // RenameWorksheetButton
            // 
            this.RenameWorksheetButton.Image = global::SubmissionCollector.Properties.Resources.Rename;
            this.RenameWorksheetButton.Label = "Rename";
            this.RenameWorksheetButton.Name = "RenameWorksheetButton";
            this.RenameWorksheetButton.ShowImage = true;
            this.RenameWorksheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RenameWorksheetButton_Click);
            // 
            // MoveWorksheetLeftButtton
            // 
            this.MoveWorksheetLeftButtton.Image = global::SubmissionCollector.Properties.Resources.LeftArrow;
            this.MoveWorksheetLeftButtton.Label = "Move Left";
            this.MoveWorksheetLeftButtton.Name = "MoveWorksheetLeftButtton";
            this.MoveWorksheetLeftButtton.ShowImage = true;
            this.MoveWorksheetLeftButtton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveWorksheetLeftButton_Click);
            // 
            // MoveWorksheetRightButtton
            // 
            this.MoveWorksheetRightButtton.Image = global::SubmissionCollector.Properties.Resources.RightArrow;
            this.MoveWorksheetRightButtton.Label = "Move Right";
            this.MoveWorksheetRightButtton.Name = "MoveWorksheetRightButtton";
            this.MoveWorksheetRightButtton.ShowImage = true;
            this.MoveWorksheetRightButtton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveWorksheetRightButton_Click);
            // 
            // FactorsButton
            // 
            this.FactorsButton.Label = "Add Factor Worksheet";
            this.FactorsButton.Name = "FactorsButton";
            this.FactorsButton.ScreenTip = "Creates custom worksheets containing OLFs and LDFs";
            this.FactorsButton.ShowImage = true;
            this.FactorsButton.SuperTip = "Creates custom worksheets based on the submsision segments and historical periods" +
    " already entered.";
            this.FactorsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FactorsButton_Click);
            // 
            // UploadSplitButton
            // 
            this.UploadSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UploadSplitButton.Image = global::SubmissionCollector.Properties.Resources.CloudChecked;
            this.UploadSplitButton.Items.Add(this.UploadButton);
            this.UploadSplitButton.Items.Add(this.UploadPreviewButton);
            this.UploadSplitButton.KeyTip = "TL";
            this.UploadSplitButton.Label = "Upload";
            this.UploadSplitButton.Name = "UploadSplitButton";
            this.UploadSplitButton.ScreenTip = "Upload (and save) workbook content into global system";
            this.UploadSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoadSplitButton_Click);
            // 
            // UploadButton
            // 
            this.UploadButton.Image = global::SubmissionCollector.Properties.Resources.CloudChecked;
            this.UploadButton.Label = "Upload";
            this.UploadButton.Name = "UploadButton";
            this.UploadButton.ScreenTip = "Upload (and save) workbook content into global system";
            this.UploadButton.ShowImage = true;
            this.UploadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UploadButton_Click);
            // 
            // UploadPreviewButton
            // 
            this.UploadPreviewButton.Image = global::SubmissionCollector.Properties.Resources.Glasses;
            this.UploadPreviewButton.Label = "Upload Preview";
            this.UploadPreviewButton.Name = "UploadPreviewButton";
            this.UploadPreviewButton.ShowImage = true;
            this.UploadPreviewButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UploadPreviewButton_Click);
            // 
            // DeleteFromDatabaseButton
            // 
            this.DeleteFromDatabaseButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DeleteFromDatabaseButton.Image = global::SubmissionCollector.Properties.Resources.DeleteX;
            this.DeleteFromDatabaseButton.KeyTip = "TD";
            this.DeleteFromDatabaseButton.Label = "Delete";
            this.DeleteFromDatabaseButton.Name = "DeleteFromDatabaseButton";
            this.DeleteFromDatabaseButton.ScreenTip = "Delete database package";
            this.DeleteFromDatabaseButton.ShowImage = true;
            this.DeleteFromDatabaseButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeletePackageButton_Click);
            // 
            // ShowRatingAnalysesButton
            // 
            this.ShowRatingAnalysesButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShowRatingAnalysesButton.Image = global::SubmissionCollector.Properties.Resources.BulletedList;
            this.ShowRatingAnalysesButton.KeyTip = "TB";
            this.ShowRatingAnalysesButton.Label = "BEX";
            this.ShowRatingAnalysesButton.Name = "ShowRatingAnalysesButton";
            this.ShowRatingAnalysesButton.ScreenTip = "Show BEX Rating Analyses Currently Using this Package";
            this.ShowRatingAnalysesButton.ShowImage = true;
            this.ShowRatingAnalysesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowRatingAnalysisButton_Click);
            // 
            // ServerHistoryButton
            // 
            this.ServerHistoryButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ServerHistoryButton.Image = global::SubmissionCollector.Properties.Resources.Parchment;
            this.ServerHistoryButton.KeyTip = "TH";
            this.ServerHistoryButton.Label = "History";
            this.ServerHistoryButton.Name = "ServerHistoryButton";
            this.ServerHistoryButton.ScreenTip = "Show global system communication history";
            this.ServerHistoryButton.ShowImage = true;
            this.ServerHistoryButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ServerHistoryButton2_Click);
            // 
            // DecouplePackageSplitButton
            // 
            this.DecouplePackageSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DecouplePackageSplitButton.Image = global::SubmissionCollector.Properties.Resources.SplitFiles;
            this.DecouplePackageSplitButton.Items.Add(this.DecouplePackageButton);
            this.DecouplePackageSplitButton.Items.Add(this.DecoupleSegmentButton);
            this.DecouplePackageSplitButton.Items.Add(this.DecoupleComponentButton);
            this.DecouplePackageSplitButton.KeyTip = "TU";
            this.DecouplePackageSplitButton.Label = "Decouple Package";
            this.DecouplePackageSplitButton.Name = "DecouplePackageSplitButton";
            this.DecouplePackageSplitButton.ScreenTip = "Decouple database package from workbook";
            this.DecouplePackageSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DecouplePackageSplitButton_Click);
            // 
            // DecouplePackageButton
            // 
            this.DecouplePackageButton.Label = "Decouple package";
            this.DecouplePackageButton.Name = "DecouplePackageButton";
            this.DecouplePackageButton.ScreenTip = "Decouple database package from workbook";
            this.DecouplePackageButton.ShowImage = true;
            this.DecouplePackageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DecouplePackageButton_Click);
            // 
            // DecoupleSegmentButton
            // 
            this.DecoupleSegmentButton.Label = "Decouple submission segment";
            this.DecoupleSegmentButton.Name = "DecoupleSegmentButton";
            this.DecoupleSegmentButton.ScreenTip = "Decouple database submission segment from workbook";
            this.DecoupleSegmentButton.ShowImage = true;
            this.DecoupleSegmentButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DecoupleSegmentButtonClick);
            // 
            // DecoupleComponentButton
            // 
            this.DecoupleComponentButton.Label = "Decouple single component";
            this.DecoupleComponentButton.Name = "DecoupleComponentButton";
            this.DecoupleComponentButton.ScreenTip = "Decouple database component from workbook";
            this.DecoupleComponentButton.ShowImage = true;
            this.DecoupleComponentButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DecoupleComponentButton_Click);
            // 
            // ValidateButton
            // 
            this.ValidateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ValidateButton.Image = global::SubmissionCollector.Properties.Resources.CheckFile;
            this.ValidateButton.KeyTip = "VV";
            this.ValidateButton.Label = "Validate";
            this.ValidateButton.Name = "ValidateButton";
            this.ValidateButton.ScreenTip = "Validate workbook content";
            this.ValidateButton.ShowImage = true;
            this.ValidateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ValidateButton_Click);
            // 
            // QualityControlSplitButton
            // 
            this.QualityControlSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.QualityControlSplitButton.Image = global::SubmissionCollector.Properties.Resources.ThumbsUp;
            this.QualityControlSplitButton.Items.Add(this.QualityControlButton);
            this.QualityControlSplitButton.Items.Add(this.QualityControlDocumentationButton);
            this.QualityControlSplitButton.KeyTip = "Q";
            this.QualityControlSplitButton.Label = "QC";
            this.QualityControlSplitButton.Name = "QualityControlSplitButton";
            this.QualityControlSplitButton.ScreenTip = "Quality Control";
            this.QualityControlSplitButton.SuperTip = "Identifies potential data inconsistencies";
            this.QualityControlSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QualityControlSplitButton_Click);
            // 
            // QualityControlButton
            // 
            this.QualityControlButton.Image = global::SubmissionCollector.Properties.Resources.ThumbsUp;
            this.QualityControlButton.Label = "Quality Control";
            this.QualityControlButton.Name = "QualityControlButton";
            this.QualityControlButton.ScreenTip = "Quality Control";
            this.QualityControlButton.ShowImage = true;
            this.QualityControlButton.SuperTip = "Identifies potential data inconsistencies";
            this.QualityControlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QualityControlButton_Click);
            // 
            // QualityControlDocumentationButton
            // 
            this.QualityControlDocumentationButton.Image = global::SubmissionCollector.Properties.Resources.Document;
            this.QualityControlDocumentationButton.Label = "QC Documentation";
            this.QualityControlDocumentationButton.Name = "QualityControlDocumentationButton";
            this.QualityControlDocumentationButton.ShowImage = true;
            this.QualityControlDocumentationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.QualityControlDocumentationButton_Click);
            // 
            // RenewSplitButton
            // 
            this.RenewSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RenewSplitButton.Image = global::SubmissionCollector.Properties.Resources.Renew;
            this.RenewSplitButton.Items.Add(this.RenewButton);
            this.RenewSplitButton.Items.Add(this.DeleteRenewProperties);
            this.RenewSplitButton.Items.Add(this.RenewHelpButton);
            this.RenewSplitButton.KeyTip = "N";
            this.RenewSplitButton.Label = "Renew";
            this.RenewSplitButton.Name = "RenewSplitButton";
            this.RenewSplitButton.ScreenTip = "Renew Workbook";
            this.RenewSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RenewSplitButton_Click);
            // 
            // RenewButton
            // 
            this.RenewButton.Image = global::SubmissionCollector.Properties.Resources.Renew;
            this.RenewButton.Label = "Renew";
            this.RenewButton.Name = "RenewButton";
            this.RenewButton.ScreenTip = "Renew Workbook";
            this.RenewButton.ShowImage = true;
            this.RenewButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RenewButton_Click);
            // 
            // DeleteRenewProperties
            // 
            this.DeleteRenewProperties.Image = global::SubmissionCollector.Properties.Resources.DeleteX;
            this.DeleteRenewProperties.Label = "Delete Renew From";
            this.DeleteRenewProperties.Name = "DeleteRenewProperties";
            this.DeleteRenewProperties.ScreenTip = "Deletes renewed from workbook full name and time stamp";
            this.DeleteRenewProperties.ShowImage = true;
            this.DeleteRenewProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteRenewProperties_Click);
            // 
            // RenewHelpButton
            // 
            this.RenewHelpButton.Image = global::SubmissionCollector.Properties.Resources.Document;
            this.RenewHelpButton.Label = "Renew Documentation";
            this.RenewHelpButton.Name = "RenewHelpButton";
            this.RenewHelpButton.ShowImage = true;
            this.RenewHelpButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RenewHelpButton_Click);
            // 
            // WorkercCompMenu
            // 
            this.WorkercCompMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.WorkercCompMenu.Image = global::SubmissionCollector.Properties.Resources.WC;
            this.WorkercCompMenu.Items.Add(this.WorkersCompConfigureMenu);
            this.WorkercCompMenu.Items.Add(this.WorkerCompClassCodeMappingButton);
            this.WorkercCompMenu.Items.Add(this.WorkersCompLookUpMenu);
            this.WorkercCompMenu.KeyTip = "W";
            this.WorkercCompMenu.Label = "Specific";
            this.WorkercCompMenu.Name = "WorkercCompMenu";
            this.WorkercCompMenu.ShowImage = true;
            // 
            // WorkersCompConfigureMenu
            // 
            this.WorkersCompConfigureMenu.Image = global::SubmissionCollector.Properties.Resources.Arrange;
            this.WorkersCompConfigureMenu.Items.Add(this.ChangeWorkersCompClassCodeButton);
            this.WorkersCompConfigureMenu.Items.Add(this.ChangeWorkersCompStateByHazardGroupButton);
            this.WorkersCompConfigureMenu.Label = "Reconfigure";
            this.WorkersCompConfigureMenu.Name = "WorkersCompConfigureMenu";
            this.WorkersCompConfigureMenu.ShowImage = true;
            // 
            // ChangeWorkersCompClassCodeButton
            // 
            this.ChangeWorkersCompClassCodeButton.Label = "Toggle Class Code Usage";
            this.ChangeWorkersCompClassCodeButton.Name = "ChangeWorkersCompClassCodeButton";
            this.ChangeWorkersCompClassCodeButton.ScreenTip = "Toggle between using class codes and using state by hazard profile";
            this.ChangeWorkersCompClassCodeButton.ShowImage = true;
            this.ChangeWorkersCompClassCodeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeWorkersCompClassCodeButton_Click);
            // 
            // ChangeWorkersCompStateByHazardGroupButton
            // 
            this.ChangeWorkersCompStateByHazardGroupButton.Label = "Toggle State By Hazard Group Independence";
            this.ChangeWorkersCompStateByHazardGroupButton.Name = "ChangeWorkersCompStateByHazardGroupButton";
            this.ChangeWorkersCompStateByHazardGroupButton.ScreenTip = "Toggle between populating state/hazard group and populating state and hazard grou" +
    "p independently";
            this.ChangeWorkersCompStateByHazardGroupButton.ShowImage = true;
            this.ChangeWorkersCompStateByHazardGroupButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeWorkersCompStateByHazardGroupButton_Click);
            // 
            // WorkerCompClassCodeMappingButton
            // 
            this.WorkerCompClassCodeMappingButton.Image = global::SubmissionCollector.Properties.Resources.Map;
            this.WorkerCompClassCodeMappingButton.Label = "Map";
            this.WorkerCompClassCodeMappingButton.Name = "WorkerCompClassCodeMappingButton";
            this.WorkerCompClassCodeMappingButton.ScreenTip = "Map Class Codes to Hazard Groups";
            this.WorkerCompClassCodeMappingButton.ShowImage = true;
            this.WorkerCompClassCodeMappingButton.Tag = "";
            this.WorkerCompClassCodeMappingButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WorkerCompClassCodeMappingButton_Click);
            // 
            // WorkersCompLookUpMenu
            // 
            this.WorkersCompLookUpMenu.Image = global::SubmissionCollector.Properties.Resources.DatabaseView;
            this.WorkersCompLookUpMenu.Items.Add(this.ShowWorkersCompClassCodesButton);
            this.WorkersCompLookUpMenu.Items.Add(this.ShowWorkersCompClassCodesQueryButton);
            this.WorkersCompLookUpMenu.Label = "Class Code Reference Data";
            this.WorkersCompLookUpMenu.Name = "WorkersCompLookUpMenu";
            this.WorkersCompLookUpMenu.ShowImage = true;
            // 
            // ShowWorkersCompClassCodesButton
            // 
            this.ShowWorkersCompClassCodesButton.Label = "View By State";
            this.ShowWorkersCompClassCodesButton.Name = "ShowWorkersCompClassCodesButton";
            this.ShowWorkersCompClassCodesButton.ScreenTip = "View Reference Data Class Codes By State";
            this.ShowWorkersCompClassCodesButton.ShowImage = true;
            this.ShowWorkersCompClassCodesButton.Tag = "";
            this.ShowWorkersCompClassCodesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowWorkersCompClassCodesButton_Click);
            // 
            // ShowWorkersCompClassCodesQueryButton
            // 
            this.ShowWorkersCompClassCodesQueryButton.Label = "Search";
            this.ShowWorkersCompClassCodesQueryButton.Name = "ShowWorkersCompClassCodesQueryButton";
            this.ShowWorkersCompClassCodesQueryButton.ScreenTip = "Search Reference Data Class Codes";
            this.ShowWorkersCompClassCodesQueryButton.ShowImage = true;
            this.ShowWorkersCompClassCodesQueryButton.Tag = "";
            this.ShowWorkersCompClassCodesQueryButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowWorkersCompClassCodesQueryButton_Click);
            // 
            // DevelopmentToolsMenu
            // 
            this.DevelopmentToolsMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DevelopmentToolsMenu.Image = global::SubmissionCollector.Properties.Resources.DevTools;
            this.DevelopmentToolsMenu.Items.Add(this.ShowServerPackageAsJsonButton);
            this.DevelopmentToolsMenu.Items.Add(this.CreateJsonButton);
            this.DevelopmentToolsMenu.Items.Add(this.button1);
            this.DevelopmentToolsMenu.Items.Add(this.ShowUserPrefsAsJsonButton);
            this.DevelopmentToolsMenu.Items.Add(this.ShowIsDirtySplitButton);
            this.DevelopmentToolsMenu.Items.Add(this.LedgerButton);
            this.DevelopmentToolsMenu.Items.Add(this.ShowRangeNamesButton);
            this.DevelopmentToolsMenu.Items.Add(this.HideRangeNamesButton);
            this.DevelopmentToolsMenu.Items.Add(this.FixCheckboxSynchronizationButton);
            this.DevelopmentToolsMenu.KeyTip = "DT";
            this.DevelopmentToolsMenu.Label = "Tools";
            this.DevelopmentToolsMenu.Name = "DevelopmentToolsMenu";
            this.DevelopmentToolsMenu.ShowImage = true;
            // 
            // ShowServerPackageAsJsonButton
            // 
            this.ShowServerPackageAsJsonButton.Image = global::SubmissionCollector.Properties.Resources.Json;
            this.ShowServerPackageAsJsonButton.Label = "Get Server Package";
            this.ShowServerPackageAsJsonButton.Name = "ShowServerPackageAsJsonButton";
            this.ShowServerPackageAsJsonButton.ShowImage = true;
            this.ShowServerPackageAsJsonButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowServerPackageAsJsonButton_Click);
            // 
            // CreateJsonButton
            // 
            this.CreateJsonButton.Image = global::SubmissionCollector.Properties.Resources.Json;
            this.CreateJsonButton.Label = "Create Package Json";
            this.CreateJsonButton.Name = "CreateJsonButton";
            this.CreateJsonButton.ShowImage = true;
            this.CreateJsonButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateJsonButton_Click);
            // 
            // button1
            // 
            this.button1.Image = global::SubmissionCollector.Properties.Resources.Xml;
            this.button1.Label = "Show XML Part";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowCustomXmlPartButton_Click);
            // 
            // ShowUserPrefsAsJsonButton
            // 
            this.ShowUserPrefsAsJsonButton.Image = global::SubmissionCollector.Properties.Resources.Json;
            this.ShowUserPrefsAsJsonButton.Label = "Show User Prefs Json";
            this.ShowUserPrefsAsJsonButton.Name = "ShowUserPrefsAsJsonButton";
            this.ShowUserPrefsAsJsonButton.ShowImage = true;
            this.ShowUserPrefsAsJsonButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowUserPreferencesAsJsonButton_Click);
            // 
            // ShowIsDirtySplitButton
            // 
            this.ShowIsDirtySplitButton.Image = global::SubmissionCollector.Properties.Resources.CloudChecked;
            this.ShowIsDirtySplitButton.Items.Add(this.ShoeIsDirtyButton);
            this.ShowIsDirtySplitButton.Items.Add(this.SetPackageIsDirtyTrueButton);
            this.ShowIsDirtySplitButton.Items.Add(this.SetPackageIsDirtyFalseButton);
            this.ShowIsDirtySplitButton.Label = "Sync Status";
            this.ShowIsDirtySplitButton.Name = "ShowIsDirtySplitButton";
            this.ShowIsDirtySplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowDirtyFlagsSplitButton_Click);
            // 
            // ShoeIsDirtyButton
            // 
            this.ShoeIsDirtyButton.Label = "Show Sync Status";
            this.ShoeIsDirtyButton.Name = "ShoeIsDirtyButton";
            this.ShoeIsDirtyButton.ShowImage = true;
            this.ShoeIsDirtyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowIsDirtyFlagsButton_Click);
            // 
            // SetPackageIsDirtyTrueButton
            // 
            this.SetPackageIsDirtyTrueButton.Label = "Set All to Out of Sync";
            this.SetPackageIsDirtyTrueButton.Name = "SetPackageIsDirtyTrueButton";
            this.SetPackageIsDirtyTrueButton.ShowImage = true;
            this.SetPackageIsDirtyTrueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetPackageIsDirtyTrueButton_Click);
            // 
            // SetPackageIsDirtyFalseButton
            // 
            this.SetPackageIsDirtyFalseButton.Label = "Set All To In Sync";
            this.SetPackageIsDirtyFalseButton.Name = "SetPackageIsDirtyFalseButton";
            this.SetPackageIsDirtyFalseButton.ShowImage = true;
            this.SetPackageIsDirtyFalseButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetPackageIsDirtyFalseButton_Click);
            // 
            // LedgerButton
            // 
            this.LedgerButton.Image = global::SubmissionCollector.Properties.Resources.Book;
            this.LedgerButton.Label = "Ledger";
            this.LedgerButton.Name = "LedgerButton";
            this.LedgerButton.ShowImage = true;
            this.LedgerButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LedgerSourceIdButton_Click);
            // 
            // ShowRangeNamesButton
            // 
            this.ShowRangeNamesButton.Label = "Show Range Names";
            this.ShowRangeNamesButton.Name = "ShowRangeNamesButton";
            this.ShowRangeNamesButton.ShowImage = true;
            this.ShowRangeNamesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowRangeNamesButton_Click);
            // 
            // HideRangeNamesButton
            // 
            this.HideRangeNamesButton.Label = "Hide Range Names";
            this.HideRangeNamesButton.Name = "HideRangeNamesButton";
            this.HideRangeNamesButton.ShowImage = true;
            this.HideRangeNamesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HideRangeNamesButton_Click);
            // 
            // FixCheckboxSynchronizationButton
            // 
            this.FixCheckboxSynchronizationButton.Label = "Fix Checkbox Synchronization";
            this.FixCheckboxSynchronizationButton.Name = "FixCheckboxSynchronizationButton";
            this.FixCheckboxSynchronizationButton.ShowImage = true;
            this.FixCheckboxSynchronizationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FixCheckboxSynchronizationButton_Click);
            // 
            // AppDataButton
            // 
            this.AppDataButton.Image = global::SubmissionCollector.Properties.Resources.Folder;
            this.AppDataButton.KeyTip = "DF";
            this.AppDataButton.Label = "App Data Folder";
            this.AppDataButton.Name = "AppDataButton";
            this.AppDataButton.ShowImage = true;
            this.AppDataButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AppDataButton_Click);
            // 
            // RebuildWorkbookButton
            // 
            this.RebuildWorkbookButton.Image = global::SubmissionCollector.Properties.Resources.Puzzle;
            this.RebuildWorkbookButton.KeyTip = "DR";
            this.RebuildWorkbookButton.Label = "Rebuild Workbook";
            this.RebuildWorkbookButton.Name = "RebuildWorkbookButton";
            this.RebuildWorkbookButton.ScreenTip = "Rebuild workbook from submission data package when prevoius workbook is lost";
            this.RebuildWorkbookButton.ShowImage = true;
            this.RebuildWorkbookButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RebuildButton_Click);
            // 
            // UserPreferencesSplitButton
            // 
            this.UserPreferencesSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UserPreferencesSplitButton.Image = global::SubmissionCollector.Properties.Resources.Gear;
            this.UserPreferencesSplitButton.Items.Add(this.UserPreferencesButton);
            this.UserPreferencesSplitButton.Items.Add(this.UserPreferencesResetButton);
            this.UserPreferencesSplitButton.Items.Add(this.RefreshReferenceDataButton);
            this.UserPreferencesSplitButton.KeyTip = "P";
            this.UserPreferencesSplitButton.Label = "User Preferences";
            this.UserPreferencesSplitButton.Name = "UserPreferencesSplitButton";
            this.UserPreferencesSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UserPreferencesSplitButton_Click);
            // 
            // UserPreferencesButton
            // 
            this.UserPreferencesButton.Image = global::SubmissionCollector.Properties.Resources.Gear;
            this.UserPreferencesButton.Label = "Edit User Preferences";
            this.UserPreferencesButton.Name = "UserPreferencesButton";
            this.UserPreferencesButton.ShowImage = true;
            this.UserPreferencesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UserPreferencesButton_Click);
            // 
            // UserPreferencesResetButton
            // 
            this.UserPreferencesResetButton.Image = global::SubmissionCollector.Properties.Resources.Reset;
            this.UserPreferencesResetButton.Label = "Reset User Preferences";
            this.UserPreferencesResetButton.Name = "UserPreferencesResetButton";
            this.UserPreferencesResetButton.ScreenTip = "Reset user preferences to defaults";
            this.UserPreferencesResetButton.ShowImage = true;
            this.UserPreferencesResetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UserPreferencesResetButton_Click);
            // 
            // RefreshReferenceDataButton
            // 
            this.RefreshReferenceDataButton.Image = global::SubmissionCollector.Properties.Resources.Reload;
            this.RefreshReferenceDataButton.Label = "Reload Reference Data";
            this.RefreshReferenceDataButton.Name = "RefreshReferenceDataButton";
            this.RefreshReferenceDataButton.ShowImage = true;
            this.RefreshReferenceDataButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RefreshReferenceDataButton_Click);
            // 
            // AboutMenu
            // 
            this.AboutMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AboutMenu.Image = global::SubmissionCollector.Properties.Resources.Information;
            this.AboutMenu.Items.Add(this.GetStartedButton);
            this.AboutMenu.Items.Add(this.BuildVersionButton);
            this.AboutMenu.Items.Add(this.IsCompatibleButton);
            this.AboutMenu.Items.Add(this.GetUrlsButton);
            this.AboutMenu.KeyTip = "A";
            this.AboutMenu.Label = "About";
            this.AboutMenu.Name = "AboutMenu";
            this.AboutMenu.ShowImage = true;
            // 
            // GetStartedButton
            // 
            this.GetStartedButton.Image = global::SubmissionCollector.Properties.Resources.StartingFlag;
            this.GetStartedButton.Label = "Get Started";
            this.GetStartedButton.Name = "GetStartedButton";
            this.GetStartedButton.ShowImage = true;
            this.GetStartedButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetStartedButton_Click);
            // 
            // BuildVersionButton
            // 
            this.BuildVersionButton.Image = global::SubmissionCollector.Properties.Resources.CircledV;
            this.BuildVersionButton.Label = "Deployment Detail";
            this.BuildVersionButton.Name = "BuildVersionButton";
            this.BuildVersionButton.ShowImage = true;
            this.BuildVersionButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BuildVersionButton_Click);
            // 
            // IsCompatibleButton
            // 
            this.IsCompatibleButton.Image = global::SubmissionCollector.Properties.Resources.CircledV;
            this.IsCompatibleButton.Label = "Workbook Version / Compatibility";
            this.IsCompatibleButton.Name = "IsCompatibleButton";
            this.IsCompatibleButton.ScreenTip = "Build information describing the supporting files";
            this.IsCompatibleButton.ShowImage = true;
            this.IsCompatibleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.IsCompatibleButton_Click);
            // 
            // GetUrlsButton
            // 
            this.GetUrlsButton.Image = global::SubmissionCollector.Properties.Resources.Url;
            this.GetUrlsButton.Label = "URLs";
            this.GetUrlsButton.Name = "GetUrlsButton";
            this.GetUrlsButton.ShowImage = true;
            this.GetUrlsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetUrlsButton_Click);
            // 
            // button2
            // 
            this.button2.Image = global::SubmissionCollector.Properties.Resources.Json;
            this.button2.Label = "Show User Prefs Json";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowUserPreferencesAsJsonButton_Click);
            // 
            // SubmissionRibbon
            // 
            this.Name = "SubmissionRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.SubmissionTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SubmissionRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.SubmissionTab.ResumeLayout(false);
            this.SubmissionTab.PerformLayout();
            this.InventoryGroup.ResumeLayout(false);
            this.InventoryGroup.PerformLayout();
            this.SegmentGroup.ResumeLayout(false);
            this.SegmentGroup.PerformLayout();
            this.RangeUtilitiesGroup.ResumeLayout(false);
            this.RangeUtilitiesGroup.PerformLayout();
            this.UtilitiesGroup.ResumeLayout(false);
            this.UtilitiesGroup.PerformLayout();
            this.DatabaseGroup.ResumeLayout(false);
            this.DatabaseGroup.PerformLayout();
            this.WorkbookGroup.ResumeLayout(false);
            this.WorkbookGroup.PerformLayout();
            this.AdminGroup.ResumeLayout(false);
            this.AdminGroup.PerformLayout();
            this.UserPreferencesButtonGroup.ResumeLayout(false);
            this.UserPreferencesButtonGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab SubmissionTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DuplicateRiskClassButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TransposeMatrixButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DatabaseGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateJsonButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RangeUtilitiesGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertRowsIntoRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteRowsFromRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup UtilitiesGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RefreshReferenceDataButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton DecouplePackageSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DecouplePackageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowServerPackageAsJsonButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddWorksheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteWorksheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SortByStateNameButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SortByStateIdButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SortByStateIdWithCwOnTopButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SortByStateNameWithCwOnTopButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetUrlsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RenameWorksheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MoveWorksheetRightButtton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MoveWorksheetLeftButtton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetPackageIsDirtyFalseButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton PaneToggleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertColumnsIntoRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteColumnsFromRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton InsertRowsIntoRangeSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SegmentGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddSegmentButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteSegmentButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton AddSegmentSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu StateSorterMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton QuickFillDatesWithValuesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SublineWizardButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShoeIsDirtyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetPackageIsDirtyTrueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton ShowIsDirtySplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup UserPreferencesButtonGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UserPreferencesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DecoupleComponentButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteFromDatabaseButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DecoupleSegmentButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetStartedButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UserPreferencesResetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UmbrellaPolicyProfileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MoveRangeRightButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MoveRangeLeftButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BuildVersionButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UploadPreviewButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu MoveComponentsMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu DevelopmentToolsMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu CustomWorksheetMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu ComponentArrangeMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowUserPrefsAsJsonButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ValidateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LedgerButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeTivRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton RenewSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RenewButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RenewHelpButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton QualityControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteRenewProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton QualityControlSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton QualityControlDocumentationButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton UserPreferencesSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu AboutMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup InventoryGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup WorkbookGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton UploadSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton IsCompatibleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ServerHistoryButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton QuickFillDatesWithValuesSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton QuickFillDatesWithFormulasButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AdminGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertRowsIntoRangeAlternativeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowRatingAnalysesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FactorsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReformatterButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EstimateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton FormatSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UmbrellaSelectorButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton UmbrellaSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AppDataButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowRangeNamesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton HideRangeNamesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReplaceButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FixCheckboxSynchronizationButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RebuildWorkbookButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu PolicyProfileDimensionMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeWorkersCompClassCodeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WorkerCompClassCodeMappingButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowWorkersCompClassCodesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowWorkersCompClassCodesQueryButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu WorkercCompMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeWorkersCompStateByHazardGroupButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu WorkersCompConfigureMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu WorkersCompLookUpMenu;
    }

    partial class ThisRibbonCollection
    {
        internal SubmissionRibbon SubmissionRibbon
        {
            get { return this.GetRibbon<SubmissionRibbon>(); }
        }
    }
}
