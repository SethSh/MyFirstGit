using System.Linq;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class WorkbookRebuilderManager
    {
        private readonly WorkbookRebuilder _rebuilder;
        
        public WorkbookRebuilderManager(WorkbookRebuilder rebuilder)
        {
            _rebuilder = rebuilder;
        }

        internal void Manage(IPackage package)
        {
            if (package.HasSourceId)
            {
                var sourceIdMessage = "Rebuild Workbook requires that this workbook isn't already couple with a " +
                                      $"{BexConstants.ServerDatabaseName.ToLower()} {BexConstants.PackageName.ToLower()}";
                MessageHelper.Show(sourceIdMessage, MessageType.Stop);
                return;
            }

            if (package.Segments.Any())
            {
                var segmentsMessage = $"Rebuild Workbook requires that this workbook contains no {BexConstants.SegmentName}s";
                MessageHelper.Show(segmentsMessage, MessageType.Stop);
                return;
            }

            var serverPackageIdAsString = MessageHelper.ShowInputBox("Enter data package ID", "0");
            if (!long.TryParse(serverPackageIdAsString, out var serverPackageId))
            {
                MessageHelper.Show("Not a number", MessageType.Stop);
                return;
            }
            
            _rebuilder.RebuildWithProgress(package, serverPackageId);
        }
    }
}
