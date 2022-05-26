using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class DecoupleManager
    {
        public void DecoupleSegmentSelected(IWorkbookLogger logger)
        {
            try
            {
                DecoupleSegmentSelected();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"{BexConstants.DecoupleName.ToStartOfSentence()} {BexConstants.SegmentName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public void DecoupleComponent(IWorkbookLogger logger)
        {
            try
            {
                DecoupleComponent();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"{BexConstants.DecoupleName.ToStartOfSentence()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static void ClearLedger(IExcelComponent excelComponent)
        {
            var exposureSet = excelComponent as ExposureSet;
            exposureSet?.Ledger.Clear();

            var aggregateLossSet = excelComponent as AggregateLossSet;
            aggregateLossSet?.Ledger.Clear();

            var individualLossSet = excelComponent as IndividualLossSet;
            individualLossSet?.Ledger.Clear();
        }

        private static void DecoupleSegmentSelected()
        {
            if (WorkbookSaveManager.IsReadOnly)
            {
                var message = $"Can't {BexConstants.DecoupleName.ToLower()} a read-only workbook";
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            var validator = new SegmentWorksheetValidator();
            if (!validator.Validate()) return;

            var segment = validator.Segment;
            if (!segment.SourceId.HasValue)
            {
                MessageHelper.Show(
                    $"{BexConstants.SegmentName.ToStartOfSentence()} <{segment.Name}> isn't found in the {BexConstants.ServerDatabaseName.ToLower()}",
                    MessageType.Stop);
                return;
            }

            var package = segment.ParentPackage;
            if (!segment.IsStructureModifiable)
            {
                var sb = new StringBuilder();
                sb.AppendLine(package.AttachedToRatingAnalysisMessage);
                sb.AppendLine();
                sb.AppendLine(
                    $"{BexConstants.DecouplingName.ToStartOfSentence()} {BexConstants.SegmentName.ToLower()} <{segment.Name}> is blocked.");
                MessageHelper.Show(sb.ToString(), MessageType.Stop);
                return;
            }

            var confirmationResponse = MessageHelper.ShowWithYesNo($"Are you sure you want to {BexConstants.DecoupleName.ToLower()} " +
                                                                   $"{BexConstants.SegmentName.ToLower()} <{segment.Name}> from the {BexConstants.ServerDatabaseName.ToLower()}?");
            if (confirmationResponse != DialogResult.Yes) return;

            var segmentName = segment.Name;
            segment.DecoupleFromServer();
            segment.ParentPackage.AddItemToBexCommunication($"{BexConstants.DecoupleName.ToStartOfSentence()}d <{segmentName}>");

            WorkbookSaveManager.Save();

            MessageHelper.Show($"{BexConstants.DecoupleName.ToStartOfSentence()}d <{segmentName}>", MessageType.Success);
        }

        private static void DecoupleComponent()
        {
            if (WorkbookSaveManager.IsReadOnly)
            {
                var message = $"Can't {BexConstants.DecoupleName.ToLower()} a read-only workbook";
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            var identifier = new SegmentExcelComponentIdentifier();
            if (!identifier.Validate()) return;

            var excelComponent =
                identifier.Segment.ExcelComponents.SingleOrDefault(x =>
                    x.CommonExcelMatrix.RangeName == identifier.ExcelMatrix.RangeName);
            if (excelComponent == null) return;

            var fullName = excelComponent.CommonExcelMatrix.FullName;
            if (!excelComponent.SourceId.HasValue)
            {
                var message =
                    $"{fullName.ToStartOfSentence()} isn't recognized as a data component in the {BexConstants.ServerDatabaseName.ToLower()}";
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            var segment = identifier.Segment;
            var package = segment.ParentPackage;
            if (!segment.IsStructureModifiable)
            {
                var sb = new StringBuilder();
                sb.AppendLine(package.AttachedToRatingAnalysisMessage);
                sb.AppendLine();
                sb.AppendLine($"{BexConstants.DecouplingName.ToStartOfSentence()} <{fullName}> is blocked.");
                MessageHelper.Show(sb.ToString(), MessageType.Stop);
                return;
            }

            var message2 =
                $"Are you sure you want to {BexConstants.DecoupleName.ToLower()} <{fullName}> from the {BexConstants.ServerDatabaseName.ToLower()}?";
            var confirmationResponse = MessageHelper.ShowWithYesNo(message2);
            if (confirmationResponse != DialogResult.Yes) return;

            var segmentName = identifier.Segment.Name;
            excelComponent.DecoupleFromServer();
            package.AddItemToBexCommunication($"{BexConstants.DecoupleName.ToStartOfSentence()}d <{segmentName}> - <{fullName}>");

            WorkbookSaveManager.Save();

            MessageHelper.Show($"{BexConstants.DecoupleName.ToStartOfSentence()} <{fullName}>", MessageType.Success);
        }
    }
}