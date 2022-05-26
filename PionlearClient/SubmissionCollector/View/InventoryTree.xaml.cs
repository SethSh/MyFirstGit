using System.Linq;
using System.Windows;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelWorkspaceFolder.SegmentMover;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for InventoryTree.xaml
    /// </summary>
    public partial class InventoryTree
    {
        public InventoryTree()
        {
            InitializeComponent();

            var viewModel = new InventoryTreeViewModel(Globals.ThisWorkbook.ThisExcelWorkspace.Package);
            Tree.DataContext = new {Package = viewModel.PackageItems};
        }
        
        private void UpButton_OnClick(object sender, RoutedEventArgs e)
        {
            var mover = new IncreaseSegmentDisplayOrder();
            var segments = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments;
            var isValid = mover.Validate(segments);
            if (!isValid) return;

            mover.Move(segments, segments.Single(s => s.IsSelected));
            RebuildSummary();
        }

        private void DownButton_OnClick(object sender, RoutedEventArgs e)
        {
            var mover = new DecreaseSegmentDisplayOrder();
            var segments = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments;
            var isValid = mover.Validate(segments);
            if (!isValid) return;

            mover.Move(segments, segments.Single(s => s.IsSelected));
            RebuildSummary();
        }

        private void RebuildSummary()
        {
            var summaryBuilder = new ProspectiveExposureSummaryBuilder();
            summaryBuilder.Build();
        }
    }
}
