using System.Windows;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Subline;
using Xceed.Wpf.Toolkit.PropertyGrid;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for InventoryPane.xaml
    /// </summary>
    public partial class InventoryPane
    {
        public InventoryPane()
        {
            InitializeComponent();

            Globals.ThisWorkbook.ThisExcelWorkspace.Package.IsExpanded = true;
            Globals.ThisWorkbook.ThisExcelWorkspace.Package.IsSelected = true; 
            
            InventoryTree.Tree.SelectedItemChanged += TreeSelectedItemChanged;
            InventoryProperties.Grid.SelectedObjectChanged += Grid_SelectedObjectChanged;
        }

        private static void Grid_SelectedObjectChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var grid = sender as PropertyGrid;
            if (grid == null) return;

            var up = UserPreferences.ReadFromFile();
            foreach (PropertyItem prop in grid.Properties)
            {
                if (prop.IsExpandable) prop.IsExpanded = up.ArePropertyNodesExpanded; 
            }
        }

        private void TreeSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var view = (IInventoryItem) e.NewValue;

            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            if (view is Package)
            {
                InventoryProperties.Grid.SelectedObject = package;
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook) return;

                Globals.ThisWorkbook.ShowPackageWorksheet();
                return;
            }

            if (view is ISegment segment)
            {
                InventoryProperties.Grid.SelectedObject = segment;
                if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook) return;

                Globals.ThisWorkbook.ShowSegmentWorksheet(segment);
                segment.WorksheetManager.UnFilterDisplay();

                ExcelSheetActivateEventManager.RefreshRibbon(segment);
                return;
            }

            if (!(view is ISubline subline)) return;

            InventoryProperties.Grid.SelectedObject = subline;
            if (Globals.ThisWorkbook.IsCurrentlyOpeningExistingWorkbook) return;
            
            segment = subline.FindParentSegment();

            using (new ExcelEventDisabler())
            {
                Globals.ThisWorkbook.ShowSegmentWorksheet(segment);
                ExcelSheetActivateEventManager.RefreshRibbon(segment);
            }

            segment.WorksheetManager.FilterDisplay(subline);
        }
    }
}
