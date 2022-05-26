using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class LedgerRenderer
    {
        public void ShowIsDirtyFlags(IPackage package, IWorkbookLogger logger)
        {
            try
            {
                ShowIsDirtyFlags(package);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("Show ledger failed", MessageType.Stop);
            }
        }

        private static void ShowIsDirtyFlags(IPackage package)
        {
            const int mainPadLength = 20;
            const int secondPadLength = 25;

            var sb = new StringBuilder();
            sb.AppendLine($"{BexConstants.PackageName} Name: {package.Name}");
            sb.AppendLine();

            foreach (var segment in package.Segments.OrderBy(x => x.DisplayOrder))
            {
                sb.AppendLine($"{BexConstants.SegmentName} Name: {segment.Name}");
                sb.AppendLine();

                sb.AppendLine();

                var excelComponents = segment.ExcelComponents.Where(ec => ec is IProvidesLedger)
                    .OrderBy(x => x.InterDisplayOrder)
                    .ThenBy(x => x.IntraDisplayOrder);
                
                foreach (var excelComponent in excelComponents)
                {
                    sb.AppendLine(excelComponent.CommonExcelMatrix.FullName);
                    sb.AppendLine("Row".PadRight(mainPadLength) +
                                  "Source Id".PadRight(mainPadLength) +
                                  "Source Timestamp".PadRight(secondPadLength) +
                                  "Is Dirty".PadRight(mainPadLength)); 
                    
                    foreach (var loss in ((IProvidesLedger)excelComponent).Ledger)
                    {
                        sb.AppendLine($"{loss.RowId.ToString().PadRight(mainPadLength)}" +
                                      $"{loss.SourceId.ToString().PadRight(mainPadLength)}" +
                                      $"{loss.SourceTimestamp.ToString().PadRight(secondPadLength)}" +
                                      $"{loss.IsDirty.ToString().PadRight(mainPadLength)}");

                    }

                    sb.AppendLine(string.Empty);
                }
            }
            ShowDisplay(sb.ToString());
        }



        private static void ShowDisplay(string message)
        {
            var title = $"{BexConstants.ApplicationName} - {BexConstants.LedgerName}";
            
            var display = new IsDirtyFlagDisplay(message);
            var form = new IsDirtyFlagsForm(display)
            {
                Height = (int)FormSizeHeight.ExtraLarge,
                Width = (int)FormSizeWidth.Medium,
                Text = title,
                StartPosition = FormStartPosition.CenterScreen,
            };
            form.ShowDialog();
        }
    }
}