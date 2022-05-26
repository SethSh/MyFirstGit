using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class FactorsWorksheetCreator
    {
        private const string OnLevelFactorShortName = "OLF";
        private const string DeveloperFactorShortName = "LDF";
        private const string OnLevelFactorsName = "On-Level Factors";
        private const string DeveloperFactorsName = "Development Factors";
        private const int DeveloperFactorsColumnWidth = 12;

        internal static void Create(IWorkbookLogger logger)
        {
            try
            {
                Create();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Create factor worksheet failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void Create()
        {
            var noSubmissionSegmentsMessage = GetNoSubmissionSegmentsMessage();
            if (!string.IsNullOrEmpty(noSubmissionSegmentsMessage))
            {
                MessageHelper.Show(noSubmissionSegmentsMessage, MessageType.Stop);
                return;
            }

            var onLevelFactorWorksheet = OnLevelFactorShortName.GetWorksheet();
            var developmentFactorWorksheet = DeveloperFactorShortName.GetWorksheet();

            var message = GetAlreadyCreatedMessage(onLevelFactorWorksheet, developmentFactorWorksheet);
            var updateFactorOption = UpdateFactorOption.None;
            if (!string.IsNullOrEmpty(message))
            {
                var augmentedMessage = $"{message}  Choose alternative:";

                var renameMessage = GetMessage("Rename", onLevelFactorWorksheet, developmentFactorWorksheet);
                var replaceMessage = GetMessage("Replace", onLevelFactorWorksheet, developmentFactorWorksheet);

                updateFactorOption = Show(augmentedMessage, renameMessage, replaceMessage);
                if (updateFactorOption == UpdateFactorOption.Cancel) return;
            }

            var segments = Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.ToList();

            using (new WorkbookUnprotector())
            {
                var thisWorksheet = HandleSheet(onLevelFactorWorksheet, updateFactorOption, OnLevelFactorShortName);
                var onLevelStartRange = thisWorksheet.Range["B4"];
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var segment in segments)
                    {
                        DrawOlf(segment, onLevelStartRange);
                        onLevelStartRange = thisWorksheet.UsedRange.GetBottomLeftCell().Offset[3, 0];
                    }
                }

                thisWorksheet = HandleSheet(developmentFactorWorksheet, updateFactorOption, DeveloperFactorShortName);
                var developmentFactorStartRange = thisWorksheet.Range["B2"];
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var segment in segments)
                    {
                        DrawLdf(segment, developmentFactorStartRange);
                        developmentFactorStartRange = thisWorksheet.UsedRange.GetBottomLeftCell().Offset[3, 0];
                    }
                }
            }
        }

        private static string GetMessage(string actionString, Worksheet onLevelFactorWorksheet, Worksheet developmentFactorWorksheet)
        {
            if (onLevelFactorWorksheet != null && developmentFactorWorksheet != null)
            {
                return
                    $"{actionString} custom worksheets <{OnLevelFactorShortName}> and <{DeveloperFactorShortName}> and create new custom worksheets.";
            }

            if (onLevelFactorWorksheet != null)
            {
                return $"{actionString} custom worksheet <{OnLevelFactorShortName}> and create new custom worksheets.";
            }

            return $"{actionString} custom worksheet <{DeveloperFactorShortName}> and create new custom worksheets.";
        }

        private static Worksheet HandleSheet(Worksheet worksheet, UpdateFactorOption option, string worksheetName)
        {
            if (worksheet != null)
            {
                switch (option)
                {
                    case UpdateFactorOption.Delete:
                    {
                        worksheet.UsedRange.Clear();
                        break;
                    }
                    case UpdateFactorOption.Rename:
                        RenameWorkbook(worksheet);
                        worksheet = null;
                        break;
                }
            }

            if (worksheet == null)
            {
                var lastSheet = Globals.ThisWorkbook.Sheets[Globals.ThisWorkbook.Sheets.Count];
                Globals.ThisWorkbook.Worksheets.Add(After: lastSheet);
                worksheet = (Worksheet) Globals.ThisWorkbook.ActiveSheet;
                worksheet.Name = worksheetName;
            }

            return worksheet;
        }

        private static void RenameWorkbook(Worksheet worksheet)
        {
            var counter = 2;
            var worksheetName = worksheet.Name;
            bool isViableName;
            do
            {
                var newWorksheetName = $"{worksheetName} {counter}";
                if (newWorksheetName.GetWorksheet() == null)
                {
                    worksheet.Name = newWorksheetName;
                    isViableName = true;
                }
                else
                {
                    isViableName = false;
                    counter++;
                }
            } while (!isViableName);
        }

        private static UpdateFactorOption Show(string message, string renameMessage, string replaceMessage)
        {
            var viewModel = new MessageBoxForFactorsViewModel(message, renameMessage, replaceMessage);
            var messageBoxForFactors = new MessageBoxForFactors(viewModel);
            var form = new MessageBoxForFactorsForm(messageBoxForFactors)
            {
                Height = (int) FormSizeHeight.Small,
                Width = (int) FormSizeWidth.Small,
                Text = BexConstants.ApplicationName,
                StartPosition = FormStartPosition.CenterScreen,
            };
            form.ShowDialog();
            return viewModel.UpdateFactorOption;
        }

        private static void DrawOlf(ISegment segment, Range startRange)
        {
            var periodsRange = segment.ExcelMatrices.First(em => em.RangeName.Contains(ExcelConstants.PeriodSetRangeName)).GetBodyRange()
                .RemoveLastColumn();
            var periodRangeRowCount = periodsRange.Rows.Count;
            var allPeriods = periodsRange.GetContent();

            var periodCount = 0;
            for (var row = 0; row < allPeriods.GetLength(0); row++)
            {
                if (allPeriods[row, 0] != null && allPeriods[row, 1] != null) periodCount++;
            }

            var periodCounter = 0;
            var periods = new object[periodCount, 2];
            for (var row = 0; row < allPeriods.GetLength(0); row++)
            {
                if (allPeriods[row, 0] == null || allPeriods[row, 1] == null) continue;

                periods[periodCounter, 0] = allPeriods[row, 0];
                periods[periodCounter, 1] = allPeriods[row, 1];
                periodCounter++;
            }

            var sublineCount = segment.Count;

            //exhibit name
            var exhibitName = new List<string> {OnLevelFactorsName};
            var destinationRange = startRange.Offset[0, 2].Resize[1, sublineCount];
            destinationRange.GetTopLeftCell().Value2 = exhibitName.ToOneByNArray();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputLabelColor();
            destinationRange.AlignCenterAcrossSelection();

            //segment name
            var segmentNames = new List<string> {segment.Name};
            destinationRange = startRange.Resize[1, 2];
            destinationRange.GetTopLeftCell().Value2 = segmentNames.ToOneByNArray();
            destinationRange.SetInputLabelColor();
            destinationRange.AlignLeft();
            destinationRange.SetBorderAroundToOrdinary();

            //dates names
            var dateNames = new List<string> {BexConstants.StartDateName, BexConstants.EndDateName};
            destinationRange = startRange.Offset[1, 0].Resize[1, 2];
            destinationRange.Value2 = dateNames.ToOneByNArray();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputLabelColor();
            destinationRange.AlignRight();

            //subline names
            var sublineNames = segment.Select(subline => subline.ShortNameWithLob).ToList();
            destinationRange = startRange.Offset[1, 2].Resize[1, sublineCount];
            destinationRange.Value2 = sublineNames.ToOneByNArray();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputLabelColor();
            destinationRange.AlignRight();

            //periods
            var periodCountToUse = periodCount > 0 ? periodCount : periodRangeRowCount;
            destinationRange = startRange.Offset[2, 0].Resize[periodCountToUse, periods.GetLength(1)];
            if (periodCount > 0) destinationRange.Value2 = periods;
            destinationRange.FormatWithDates();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputLabelColor();
            destinationRange.SetDateColumnWidth();

            //olfs
            destinationRange = startRange.Offset[2, 2].Resize[periodCountToUse, sublineCount];
            destinationRange.FormatWithFactors();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputColor();
            destinationRange.SetFactorColumnWidth();
            //destinationRange.Value2 = 1d;
        }

        private static void DrawLdf(ISegment segment, Range startRange)
        {
            var periodsRange = segment.ExcelMatrices.First(em => em.RangeName.Contains(ExcelConstants.PeriodSetRangeName)).GetBodyRange()
                .RemoveLastColumn();
            var periodRangeRowCount = periodsRange.Rows.Count;
            var allPeriods = periodsRange.GetContent();

            var periodCount = 0;
            for (var row = 0; row < allPeriods.GetLength(0); row++)
            {
                if (allPeriods[row, 0] != null && allPeriods[row, 1] != null) periodCount++;
            }

            var periodCounter = 0;
            var periods = new object[periodCount, 2];
            for (var row = 0; row < allPeriods.GetLength(0); row++)
            {
                if (allPeriods[row, 0] == null || allPeriods[row, 1] == null) continue;

                periods[periodCounter, 0] = allPeriods[row, 0];
                periods[periodCounter, 1] = allPeriods[row, 1];
                periodCounter++;
            }

            var layerCount = 10;
            var ldfType = new List<string> {"Paid", "Rptd", "Count"};
            var ldfCount = ldfType.Count;

            //exhibit name
            var exhibitName = new List<string> {DeveloperFactorsName};
            var destinationRange = startRange.Offset[0, 2].Resize[1, layerCount * ldfCount];
            destinationRange.GetTopLeftCell().Value2 = exhibitName.ToOneByNArray();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputLabelColor();
            destinationRange.AlignCenterAcrossSelection();

            //segment name
            var segmentNames = new List<string> {segment.Name};
            destinationRange = startRange.Offset[2, 0].Resize[1, 2];
            destinationRange.GetTopLeftCell().Value2 = segmentNames.ToOneByNArray();
            destinationRange.SetInputLabelColor();
            destinationRange.AlignLeft();
            destinationRange.SetBorderAroundToOrdinary();

            //lim and att
            for (var row = 0; row < layerCount; row++)
            {
                destinationRange = startRange.Offset[2, 2 + (row * ldfCount)].Resize[1, ldfCount];
                destinationRange.AppendRow().SetBorderAroundToOrdinary();

                var limitRange = destinationRange.GetTopLeftCell();
                limitRange.SetInputColor();
                limitRange.NumberFormat = FormatExtensions.WholeNumberFormat;
                //limitRange.Value2 = (row+1) * 1e6;

                var xsRange = destinationRange.GetTopLeftCell().Offset[0, 1];
                xsRange.SetInputLabelColor();
                xsRange.Value2 = "xs";
                xsRange.AlignCenter();

                var attachmentRange = destinationRange.GetTopLeftCell().Offset[0, 2];
                attachmentRange.SetInputColor();
                attachmentRange.NumberFormat = FormatExtensions.WholeNumberFormat;
                //attachmentRange.Value2 = (row + 1) * 1.5e6;
            }

            //dates names
            var dateNames = new List<string> {BexConstants.StartDateName, BexConstants.EndDateName};
            destinationRange = startRange.Offset[3, 0].Resize[1, 2];
            destinationRange.Value2 = dateNames.ToOneByNArray();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputLabelColor();
            destinationRange.AlignRight();

            //layers 
            for (var row = 0; row < layerCount; row++)
            {
                var rowPlusOne = row + 1;
                var layerName = new List<string> {$"Layer {rowPlusOne}", string.Empty, string.Empty};
                destinationRange = startRange.Offset[1, 2 + (row * ldfCount)].Resize[1, ldfCount];
                destinationRange.Value2 = layerName.ToOneByNArray();
                destinationRange.SetBorderAroundToOrdinary();
                destinationRange.SetInputLabelColor();
                destinationRange.AlignCenterAcrossSelection();
                destinationRange.SetFactorColumnWidth(DeveloperFactorsColumnWidth);
            }

            //periods
            var periodCountToUse = periodCount > 0 ? periodCount : periodRangeRowCount;
            destinationRange = startRange.Offset[4, 0].Resize[periodCountToUse, periods.GetLength(1)];
            if (periodCount > 0) destinationRange.Value2 = periods;
            destinationRange.FormatWithDates();
            destinationRange.SetBorderAroundToOrdinary();
            destinationRange.SetInputLabelColor();
            destinationRange.SetDateColumnWidth();

            //ldf types 
            var ldfTypeArray = ldfType.ToOneByNArray();
            for (var row = 0; row < layerCount; row++)
            {
                destinationRange = startRange.Offset[3, 2 + (row * ldfCount)].Resize[1, ldfCount];
                destinationRange.Value2 = ldfTypeArray;
                destinationRange.SetInputLabelColor();
                destinationRange.AlignCenter();
                //destinationRange.SetFactorColumnWidth();
            }

            //ldfs
            destinationRange = startRange.Offset[4, 2].Resize[periodCountToUse, layerCount * ldfCount];
            destinationRange.FormatWithFactors();
            for (var row = 0; row < layerCount; row++)
            {
                destinationRange.GetTopLeftCell().Offset[0, row * ldfCount].Resize[periodCountToUse, ldfCount].SetBorderAroundToOrdinary();
            }

            destinationRange.SetInputColor();
            //destinationRange.SetFactorColumnWidth();
            //destinationRange.Value2 = 1d;
        }

        private static string GetAlreadyCreatedMessage(Worksheet olf, Worksheet ldf)
        {
            if (olf == null && ldf == null) return string.Empty;

            if (olf != null && ldf != null)
            {
                return $"Custom worksheets <{OnLevelFactorShortName}> and <{DeveloperFactorShortName}> already exist.";
            }

            return olf != null
                ? $"Custom worksheet <{OnLevelFactorShortName}> already exists."
                : $"Custom worksheet <{DeveloperFactorShortName}> already exists.";
        }

        private static string GetNoSubmissionSegmentsMessage()
        {
            return Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Any()
                ? string.Empty
                : $"No {BexConstants.SegmentName.ToLower()}s";
        }
    }
}