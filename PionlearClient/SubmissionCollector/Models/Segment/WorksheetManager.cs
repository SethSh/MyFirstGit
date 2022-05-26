using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelUtilities.RangeSizeModifier;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Profiles.ExcelComponent.Helpers;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Models.Segment
{
    public class WorksheetManager
    {
        private readonly ISegment _segment;
        
        public WorksheetManager(ISegment segment)
        {
            _segment = segment;
        }


        [JsonIgnore]
        public Worksheet Worksheet { get; set; }

        public string WorksheetName { get; set; }

        public void CreateWorksheet()
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            var templateWorksheet = (Worksheet) Globals.ThisWorkbook.Sheets[ExcelConstants.SegmentTemplateWorksheetName];
            using (new WorkbookUnprotector())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    using (new ExcelEventDisabler())
                    {
                        Worksheet worksheetDirectlyToLeft;
                        if (package.Segments.Count > 1)
                        {
                            var maximumDisplayOrder = package.Segments.Where(x => x.Guid != _segment.Guid).Max(x => x.DisplayOrder);
                            var otherSegment = package.Segments.Single(x => x.DisplayOrder == maximumDisplayOrder);
                            worksheetDirectlyToLeft = otherSegment.WorksheetManager.Worksheet;
                        }
                        else
                        {
                            worksheetDirectlyToLeft = package.Worksheet;
                        }

                        Globals.ThisWorkbook.CopyWorksheet(templateWorksheet);
                        Worksheet = (Worksheet) Globals.ThisWorkbook.ActiveSheet;
                        Worksheet.Move(After: worksheetDirectlyToLeft);

                        var zoom = UserPreferences.ReadFromFile().SegmentWorksheetZoom;
                        Globals.ThisWorkbook.Application.ActiveWindow.Zoom = zoom;

                        try
                        {
                            Worksheet.UnprotectInterface();
                            Worksheet.Name = _segment.Name;
                            Worksheet.SelectFirstCell();

                            Worksheet.ResetSourceRangeNames(ExcelConstants.SegmentTemplateRangeName, $"segment{_segment.Id}");
                            SetHeader();

                            CreateRanges();

                            WriteSegmentSublineNamesIntoOriginalRange();
                            ModifySublineRanges();

                            SetHistoricalPeriods();

                            DeleteUmbrellaMatrix();

                            var prospectiveExposureBasisName =
                                ExposureBasisFromBex.GetExposureBasisName(Convert.ToInt32(_segment.ProspectiveExposureBasis));
                            SetProspectiveExposureBasis(prospectiveExposureBasisName);

                            var historicalExposureBasisName =
                                ExposureBasisFromBex.GetExposureBasisName(Convert.ToInt16(_segment.HistoricalExposureBasis));
                            SetHistoricalExposureSetsBasis(historicalExposureBasisName);

                            var historicalPeriodTypeName =
                                HistoricalPeriodTypesFromBex.GetName(Convert.ToInt16(_segment.HistoricalPeriodType));
                            SetHistoricalPeriodType(historicalPeriodTypeName);

                            if (_segment.Count == 1)
                            {
                                SetSingleSublineValueToOne();
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(ex);
                            throw;
                        }
                        finally
                        {
                            Worksheet.ProtectInterface();
                        }
                    }
                }
            }
        }

        public void CreateUmbrellaMatrix()
        {
            if (_segment.UmbrellaExcelMatrix.HeaderRangeName.ExistsInWorkbook()) return;

            var templateRange = $"{ExcelConstants.SegmentTemplateRangeName}.{ExcelConstants.UmbrellaAllocationRangeName}".GetRange();
            var columnsCount = templateRange.Columns.Count;

            var anchorRange = _segment.SublineExcelMatrix.GetHeaderRange().GetTopRightCell().Offset[0, 2];

            anchorRange.Resize[1, columnsCount + 1].InsertColumnsToRight();
            anchorRange = _segment.SublineExcelMatrix.GetHeaderRange().GetTopRightCell().Offset[0, 2];

            var destinationRange = anchorRange.Resize[templateRange.Rows.Count, columnsCount];

            templateRange.Copy(destinationRange);
            for (var column = 0; column < columnsCount + 1; column++)
            {
                destinationRange.Offset[0, column].Resize[1, 1].EntireColumn.ColumnWidth =
                    templateRange.Offset[0, column].Resize[1, 1].EntireColumn.ColumnWidth;
            }

            destinationRange.GetFirstRow().SetInvisibleRangeName(_segment.UmbrellaExcelMatrix.HeaderRangeName);
            destinationRange.GetRangeSubset(1, 1).Resize[1, 1].SetInvisibleRangeName(_segment.UmbrellaExcelMatrix.BasisRangeName);
            destinationRange.GetRangeSubset(1, 0).SetInvisibleRangeName(_segment.UmbrellaExcelMatrix.RangeName);

            WriteUmbrellaTypesIntoRange();
        }

        public void DeleteUmbrellaMatrix()
        {
            var rangeName = _segment.UmbrellaExcelMatrix.RangeName;
            rangeName.GetRange().AppendColumn().EntireColumn.Delete();

            Globals.ThisWorkbook.Names.Item(rangeName).Delete();
            Globals.ThisWorkbook.Names.Item(_segment.UmbrellaExcelMatrix.BasisRangeName).Delete();

            var headerRangeName = _segment.UmbrellaExcelMatrix.HeaderRangeName;
            if (headerRangeName.ExistsInWorkbook())
            {
                Globals.ThisWorkbook.Names.Item(headerRangeName).Delete();
            }
        }

        public void DuplicateWorksheet(ISegment sourceSegment)
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            var sourceWorksheet = sourceSegment.WorksheetManager.Worksheet;
            using (new WorkbookUnprotector())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    using (new ExcelEventDisabler())
                    {
                        try
                        {
                            var maximumDisplayOrder = package.Segments.Where(x => x.Guid != _segment.Guid).Max(x => x.DisplayOrder);
                            var otherSegment = package.GetSegmentBasedOnDisplayOrder(maximumDisplayOrder);
                            var worksheetDirectlyToLeft = otherSegment.WorksheetManager.Worksheet;

                            Globals.ThisWorkbook.CopyWorksheet(sourceWorksheet);
                            Worksheet = (Worksheet) Globals.ThisWorkbook.ActiveSheet;
                            Worksheet.Move(After: worksheetDirectlyToLeft);

                            Worksheet.Name = _segment.Name;
                            Worksheet.SelectFirstCell();

                            Worksheet.ResetSourceRangeNames($"segment{sourceSegment.Id}", $"segment{_segment.Id}");
                            SetHeader();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(ex);
                            throw;
                        }
                        finally
                        {
                            Worksheet.ProtectInterface();
                        }
                    }
                }
            }
        }

        public void UnFilterDisplay()
        {
            foreach (var excelMatrix in _segment.ExcelMatrices.OfType<MultipleOccurrenceSegmentExcelMatrix>())
            {
                excelMatrix.ShowColumns();
            }
        }

        public void FilterDisplay(ISubline subline)
        {
            foreach (var excelMatrix in _segment.ExcelMatrices.OfType<MultipleOccurrenceSegmentExcelMatrix>())
            {
                excelMatrix.SetColumnVisibility(subline);
            }
        }

        public void ModifyWorksheetWrapper()
        {
            using (new ExcelScreenUpdateDisabler())
            {
                using (new ExcelEventDisabler())
                {
                    using (new WorkbookUnprotector())
                    {
                        ModifyWorksheet();
                    }
                }
            }
        }

        public void ModifyWorksheet()
        {
            try
            {
                Worksheet.UnprotectInterface();

                ModifyRanges();

                var sublinesAllocationInRange = GetSublineMatrixContent();
                var sublineNamesFromRange = sublinesAllocationInRange.Keys.ToList();
                var sublineCodesFromRange = SublineCodesFromBex.ConvertShortNameWithLobsToCodes(sublineNamesFromRange)
                    .Select(item => item.ToString()).ToList();
                var desiredSublineCodes = _segment.Select(x => x.Code.ToString()).ToList();
                WriteSegmentSublineNamesIntoExistingRange(desiredSublineCodes, sublineCodesFromRange);

                ModifySublineRanges();

                if (!_segment.IsUmbrella) return;

                if (_segment.ContainsAnyCommercialSublines && !_segment.UmbrellaExcelMatrix.RangeName.ExistsInWorkbook())
                {
                    CreateUmbrellaMatrix();
                }

                if (!_segment.ContainsAnyCommercialSublines && _segment.UmbrellaExcelMatrix.RangeName.ExistsInWorkbook())
                {
                    DeleteUmbrellaMatrix();
                }

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
                throw;
            }
            finally
            {
                Worksheet.ProtectInterface();
            }
        }

        public void SetHistoricalPeriodType(string value)
        {
            _segment.PeriodSet.ExcelMatrix.HeaderRangeName.GetRange().GetTopLeftCell().Value = $"Historical {value}s";
        }

        public void SetHistoricalExposureSetsBasis(string value)
        {
            _segment.ExposureSets.ForEach(exposureSet => exposureSet.ExcelMatrix.GetInputLabelRange().Value = value);
        }

        public void SetProspectiveExposureBasis(string value)
        {
            _segment.ProspectiveExposureAmountExcelMatrix.GetInputLabelRange().Value = $"Prospective {value}";
        }

        public void DeleteWorksheet()
        {
            if (Worksheet == null)
            {
                throw new InvalidOperationException("Can't find segment worksheet to delete");
            }
            Globals.ThisWorkbook.DeleteWorksheet(Worksheet);

        }
        
        
        internal static int GetPeriodCount(ISegment segment)
        {
            return segment.PeriodSet.ExcelMatrix.GetBodyRange().Rows.Count;
        }
        
        internal static int GetAggregateColumnCount(IAggregateLossSetDescriptor setDescriptor)
        {
            if (setDescriptor.IsLossAndAlaeCombined)
            {
                return setDescriptor.IsPaidAvailable ? 2 : 1;
            }

            return setDescriptor.IsPaidAvailable ? 4 : 2;
        }

        internal static int GetIndividualColumnCount(IIndividualLossSetDescriptor setDescriptor)
        {
            var columnCount = 3;
            if (setDescriptor.IsLossAndAlaeCombined && setDescriptor.IsPaidAvailable)
            {
                columnCount +=2;
            }
            else if (setDescriptor.IsLossAndAlaeCombined && !setDescriptor.IsPaidAvailable)
            {
                columnCount++;
            }
            else if (!setDescriptor.IsLossAndAlaeCombined && setDescriptor.IsPaidAvailable)
            {
                columnCount += 4;
            }
            else if (!setDescriptor.IsLossAndAlaeCombined && !setDescriptor.IsPaidAvailable)
            {
                columnCount +=2;
            }

            if (setDescriptor.IsEventCodeAvailable) columnCount++;
            if (setDescriptor.IsPolicyLimitAvailable) columnCount++;
            if (setDescriptor.IsPolicyAttachmentAvailable) columnCount++;
            if (setDescriptor.IsAccidentDateAvailable) columnCount++;
            if (setDescriptor.IsPolicyDateAvailable) columnCount++;
            if (setDescriptor.IsReportDateAvailable) columnCount++;

            return columnCount;
        }

        
        private void CreateRanges()
        {
            var hazardExcelMatrixHelper = new HazardExcelMatrixHelper(_segment);
            hazardExcelMatrixHelper.CreateRanges();

            var policyExcelMatrixHelper = new PolicyExcelMatrixHelper(_segment);
            policyExcelMatrixHelper.CreateRanges();

            var stateExcelMatrixHelper = new StateExcelMatrixHelper(_segment);
            stateExcelMatrixHelper.CreateRanges();


            var constructionTypeExcelMatrixHelper = new ConstructionTypeExcelMatrixHelper(_segment);
            constructionTypeExcelMatrixHelper.CreateRanges();

            var occupancyTypeExcelMatrixHelper = new OccupancyTypeExcelMatrixHelper(_segment);
            occupancyTypeExcelMatrixHelper.CreateRanges();

            var protectionClassValueExcelMatrixHelper = new ProtectionClassExcelMatrixHelper(_segment);
            protectionClassValueExcelMatrixHelper.CreateRanges();

            var totalInsuredValueExcelMatrixHelper = new TotalInsuredValueExcelMatrixHelper(_segment);
            totalInsuredValueExcelMatrixHelper.CreateRanges();


            var minnCompStateHazardExcelMatrixHelper = new MinnesotaRetentionExcelMatrixHelper(_segment);
            minnCompStateHazardExcelMatrixHelper.CreateRange(); 
            
            var workersCompStateHazardGroupExcelMatrixHelper = new WorkersCompStateHazardGroupExcelMatrixHelper(_segment);
            workersCompStateHazardGroupExcelMatrixHelper.CreateRange();

            if (_segment.IsWorkerCompClassCodeActive)
            {
                var workersCompClassCodeExcelMatrixHelper = new WorkersCompClassCodeExcelMatrixHelper(_segment);
                workersCompClassCodeExcelMatrixHelper.CreateRange();
            }

            var workersCompStateAttachmentExcelMatrixHelper = new WorkersCompStateAttachmentExcelMatrixHelper(_segment);
            workersCompStateAttachmentExcelMatrixHelper.CreateRange();



            _segment.PeriodSet.CommonExcelMatrix.Reformat();

            var exposureSetExcelMatrixHelper = new ExposureSetExcelMatrixHelper(_segment);
            exposureSetExcelMatrixHelper.CreateRanges();

            var aggregateLossSetExcelMatrixHelper = new AggregateLossSetExcelMatrixHelper(_segment);
            aggregateLossSetExcelMatrixHelper.CreateRanges();

            var individualLossSetExcelMatrixHelper = new IndividualLossSetExcelMatrixHelper(_segment);
            individualLossSetExcelMatrixHelper.CreateRanges();

            var rateChangeSetExcelMatrixHelper = new RateChangeSetExcelMatrixHelper(_segment);
            rateChangeSetExcelMatrixHelper.CreateRanges();
        }

        private void ModifyRanges()
        {
            var hazardExcelMatrixHelper = new HazardExcelMatrixHelper(_segment);
            hazardExcelMatrixHelper.ModifyRanges();

            var policyExcelMatrixHelper = new PolicyExcelMatrixHelper(_segment);
            policyExcelMatrixHelper.ModifyRanges();

            var stateExcelMatrixHelper = new StateExcelMatrixHelper(_segment);
            stateExcelMatrixHelper.ModifyRanges();

            
            var constructionTypeExcelMatrixHelper = new ConstructionTypeExcelMatrixHelper(_segment);
            constructionTypeExcelMatrixHelper.ModifyRanges();

            var occupancyTypeExcelMatrixHelper = new OccupancyTypeExcelMatrixHelper(_segment);
            occupancyTypeExcelMatrixHelper.ModifyRanges();

            var protectionClassExcelMatrixHelper = new ProtectionClassExcelMatrixHelper(_segment);
            protectionClassExcelMatrixHelper.ModifyRanges();

            var totalInsuredValueExcelMatrixHelper = new TotalInsuredValueExcelMatrixHelper(_segment);
            totalInsuredValueExcelMatrixHelper.ModifyRanges();


            var minnesotaRetentionExcelMatrixHelper = new MinnesotaRetentionExcelMatrixHelper(_segment);
            minnesotaRetentionExcelMatrixHelper.ModifyRange(); 
            
            var workersCompStateHazardExcelMatrixHelper = new WorkersCompStateHazardGroupExcelMatrixHelper(_segment);
            workersCompStateHazardExcelMatrixHelper.ModifyRange();

            var workersCompClassCodeExcelMatrixHelper = new WorkersCompClassCodeExcelMatrixHelper(_segment);
            workersCompClassCodeExcelMatrixHelper.ModifyRange();

            var workersCompStateAttachmentExcelMatrixHelper = new WorkersCompStateAttachmentExcelMatrixHelper(_segment);
            workersCompStateAttachmentExcelMatrixHelper.ModifyRange();


            var exposureSetExcelMatrixHelper = new ExposureSetExcelMatrixHelper(_segment);
            exposureSetExcelMatrixHelper.ModifyRanges();

            var aggregateLossSetExcelMatrixHelper = new AggregateLossSetExcelMatrixHelper(_segment);
            aggregateLossSetExcelMatrixHelper.ModifyRanges();

            var individualLossSetExcelMatrixHelper = new IndividualLossSetExcelMatrixHelper(_segment);
            individualLossSetExcelMatrixHelper.ModifyRanges();

            var rateChangeSetExcelMatrixHelper = new RateChangeSetExcelMatrixHelper(_segment);
            rateChangeSetExcelMatrixHelper.ModifyRanges();
        }
        
        private void SetHistoricalPeriods()
        {
            var count = UserPreferences.ReadFromFile().HistoricalPeriodCount;
            var delta = count - UserPreferences.HistoricalPeriodCountDefault;
            if (delta == 0) return;

            IRangeResizer rangeResizer;
            if (delta > 0)
            {
                var range = _segment.PeriodSet.ExcelMatrix.GetInputRange().GetTopLeftCell().Offset[1, 0].Resize[delta, 1];
                rangeResizer = new PeriodSetRowInserter();
                rangeResizer.SetCommonProperties(_segment.PeriodSet.ExcelMatrix, range);
                rangeResizer.ModifyRange();
            }
            else
            {
                var range = _segment.PeriodSet.ExcelMatrix.GetInputRange().GetTopLeftCell().Offset[1, 0].Resize[-delta, 1];
                rangeResizer = new PeriodSetRowDeleter();
                rangeResizer.SetCommonProperties(_segment.PeriodSet.ExcelMatrix, range);
                rangeResizer.ModifyRange();
            }
        }
        
        private void WriteUmbrellaTypesIntoRange()
        {
            _segment.UmbrellaExcelMatrix.SetProfileBasisInWorksheet();

            var commercialUmbrellaTypes = UmbrellaTypesFromBex.GetCommercialTypes()
                .OrderBy(x => x.DisplayOrder)
                .ToList();
            var rowCount = commercialUmbrellaTypes.Count;
            var content = new string[rowCount, 1];
            var counter = 0;
            commercialUmbrellaTypes.ForEach(x => content[counter++, 0] = x.UmbrellaTypeName);

            var range = _segment.UmbrellaExcelMatrix.GetBodyRange();
            
            range.Offset[1, 0].Resize[range.Rows.Count - counter, range.Columns.Count].DeleteRangeUp();
            range.GetColumn(0).Value2 = content;

            _segment.UmbrellaExcelMatrix.Reformat();
        }

        private void SetHeader()
        {
            var header = new[,] { {Globals.ThisWorkbook.ThisExcelWorkspace.Package.Name}, {_segment.Name} };
            SetAdditionalHeaderRangeValues(_segment.HeaderRangeName, header);
        }

        private Dictionary<string, string> GetSublineMatrixContent()
        {
            var sublinesAtStart = new Dictionary<string, string>();
            var excelContent = _segment.SublineExcelMatrix.RangeName.GetRangeSubset(1, 0).GetContentFormulas();
            for (var i = 0; i < excelContent.GetLength(0); i++)
            {
                var name = excelContent[i, 0].ToString();
                var formula = excelContent[i, 1].ToString();
                if (string.IsNullOrEmpty(name)) continue;

                sublinesAtStart.Add(name, formula);
            }
            return sublinesAtStart;
        }

        public Dictionary<string, string> GetUmbrellaMatrixContent()
        {
            var umbrellasAtStart = new Dictionary<string, string>();
            var excelContent = _segment.UmbrellaExcelMatrix.RangeName.GetRangeSubset(1, 0).GetContentFormulas();
            for (var i = 0; i < excelContent.GetLength(0); i++)
            {
                var name = excelContent[i, 0].ToString();
                var formula = excelContent[i, 1].ToString();
                if (string.IsNullOrEmpty(name)) continue;

                umbrellasAtStart.Add(name, formula);
            }
            return umbrellasAtStart;
        }

        private void SetSingleSublineValueToOne()
        {
            _segment.SublineExcelMatrix.RangeName.GetRangeSubset(1,1).GetTopLeftCell().Value = 1;
        }

        private void WriteSegmentSublineNamesIntoOriginalRange()
        {
            _segment.SublineExcelMatrix.SetProfileBasisInWorksheet();

            var sublineNamesRange = _segment.SublineExcelMatrix.GetBodyRange().GetFirstColumn();

            var sublineNames = _segment.Select(x => x.ShortNameWithLob).ToList();
            var extraRowCount = sublineNamesRange.Rows.Count - sublineNames.Count;
            if (extraRowCount > 0) sublineNamesRange.Offset[sublineNames.Count, 0].Resize[extraRowCount, 2].DeleteRangeUp();

            sublineNamesRange.Value2 = sublineNames.ToNByOneArray();
            _segment.SublineExcelMatrix.Reformat();
        }

        private void WriteSegmentSublineNamesIntoExistingRange(IList<string> desiredSublineCodes, IList<string> sublineCodesFromRange)
        {
            if (sublineCodesFromRange.IsEqualTo(desiredSublineCodes)) return;

            var bodyRange = _segment.SublineExcelMatrix.GetBodyRange();
            
            //step 1 => clear sublines no longer used
            var counter = 0;
            foreach (var sublineCodeFromRange in sublineCodesFromRange)
            {
                if (!desiredSublineCodes.Contains(sublineCodeFromRange))
                {
                    bodyRange.Resize[1, bodyRange.Columns.Count].Offset[counter, 0].ClearContents();
                }
                counter++;
            }

            //step 2 => delete empty rows in reverse order
            for (var row = bodyRange.Rows.Count-1; row >= 0; row--)
            {
                if (bodyRange.GetTopLeftCell().Offset[row, 0].Value2 == null)
                {
                    if (bodyRange.Rows.Count > 1)
                    {
                        bodyRange.Resize[1, bodyRange.Columns.Count].Offset[row, 0].DeleteRangeUp();
                    }
                }
            }

            var currentAllocationInRange = GetSublineMatrixContent();
            var currentNamesFromRangeSortedList = new SortedList();
            foreach (var item in currentAllocationInRange.Keys) currentNamesFromRangeSortedList.Add(item, item);

            //step 3 => write new sublines into range
            var newSublineCodes = desiredSublineCodes.Except(sublineCodesFromRange).ToList();
            var newSublineNames = SublineCodesFromBex.ConvertCodesToShortNameWithLobs(newSublineCodes.Select(item => Convert.ToInt64(item)).ToList());
            foreach (var newSublineName in newSublineNames)
            {
                currentNamesFromRangeSortedList.Add(newSublineName, newSublineName);
                var rowOffset = currentNamesFromRangeSortedList.IndexOfKey(newSublineName);

                var singleRowCase = bodyRange.Rows.Count == 1 && bodyRange.GetTopLeftCell().Value2 == null;
                if (!singleRowCase)
                {
                    bodyRange.Resize[1, bodyRange.Columns.Count].Offset[rowOffset, 0].InsertRangeDown();
                    bodyRange = _segment.SublineExcelMatrix.GetBodyRange();
                }
                
                bodyRange.GetTopLeftCell().Offset[rowOffset, 0].Value2 = newSublineName;
            }

            //step 4 => correct range name and fix sum
            bodyRange.GetTopLeftCell().Offset[-1, 0].Resize[desiredSublineCodes.Count + 1, bodyRange.Columns.Count].SetInvisibleRangeName(_segment.SublineExcelMatrix.RangeName);
            var inputRange = _segment.SublineExcelMatrix.GetInputRange();
            var sumRange = _segment.SublineExcelMatrix.GetSumRange();
            sumRange.ApplySumCheck();
            sumRange.Formula = $"=Sum({inputRange.Address})";

            //step 6 => format
            _segment.SublineExcelMatrix.Reformat();
        }

        public void WriteUmbrellaNamesIntoExistingRange(IList<string> desiredCodes, IList<string> codesFromRange)
        {
            if (codesFromRange.IsEqualTo(desiredCodes)) return;

            var bodyRange = _segment.UmbrellaExcelMatrix.GetBodyRange();
            var bodyRangeContent = bodyRange.GetContent();

            
            //step 1 => clear umbrella types no longer used
            var range = bodyRange;
            foreach (var unwantedCode in codesFromRange.Except(desiredCodes))
            {
                var unwantedName = UmbrellaTypesFromBex.GetName(Convert.ToInt32(unwantedCode));
                for (var row = 0; row < bodyRangeContent.GetLength(0); row++){
                    
                    if (!bodyRangeContent[row, 0].ToString().Equals(unwantedName)) continue;
                    range.Resize[1, range.Columns.Count].Offset[row, 0].ClearContents();
                }
            }

            //step 2 => delete empty rows in reverse order
            for (var row = range.Rows.Count-1; row >= 0; row--)
            {
                if (range.GetTopLeftCell().Offset[row, 0].Value2 == null)
                {
                    if (bodyRange.Rows.Count > 1)
                    {
                        range.Resize[1, range.Columns.Count].Offset[row, 0].DeleteRangeUp();
                    }
                }
            }

            var currentAllocationInRange = GetUmbrellaMatrixContent();
            var currentNamesFromRange = currentAllocationInRange.Keys.ToList();
            var sortedList = new SortedList<long, string>();
            foreach (var name in currentNamesFromRange)
            {
                var code = UmbrellaTypesFromBex.GetCode(name);
                sortedList.Add(code, name);
            }
            
            //step 3 => write new umbrella types into range
            var newUmbrellaCodes = desiredCodes.Except(codesFromRange).ToList();
            foreach (var newUmbrellaCode in newUmbrellaCodes.Select(item => Convert.ToInt64(item)))
            {
                var newUmbrellaName = UmbrellaTypesFromBex.GetName(newUmbrellaCode);
                sortedList.Add(newUmbrellaCode, newUmbrellaName);
                var rowOffset = sortedList.IndexOfKey(newUmbrellaCode);

                var singleRowCase = bodyRange.Rows.Count == 1 && bodyRange.GetTopLeftCell().Value2 == null;
                if (!singleRowCase)
                {
                    bodyRange.GetTopLeftCell().Resize[1, range.Columns.Count].Offset[rowOffset, 0].InsertRangeDown();
                }

                //need to re-fetch range when inserting at the top
                if (rowOffset == 0)
                {
                    bodyRange = _segment.UmbrellaExcelMatrix.GetBodyRange();
                }
                bodyRange.GetTopLeftCell().Offset[rowOffset, 0].Value2 = newUmbrellaName;
            }

            //step 4 => correct range name and fix sum
            bodyRange.GetTopLeftCell().Offset[-1, 0].Resize[desiredCodes.Count + 1, bodyRange.Columns.Count].SetInvisibleRangeName(_segment.UmbrellaExcelMatrix.RangeName);
            var inputRange = _segment.UmbrellaExcelMatrix.GetInputRange();
            var sumRange = _segment.UmbrellaExcelMatrix.GetSumRange();
            sumRange.ApplySumCheck();
            sumRange.Formula = $"=Sum({inputRange.Address})";

            //step 5 => format
            _segment.UmbrellaExcelMatrix.Reformat();
        }

        private void ModifySublineRanges()
        {
           var multipleSegmentExcelMatrices = _segment.ExcelMatrices.OfType<MultipleOccurrenceSegmentExcelMatrix>().ToList();
            var maximumSublineCount = multipleSegmentExcelMatrices.Max(excelMatrix => excelMatrix.Count);
            foreach (var excelMatrix in multipleSegmentExcelMatrices) 
            {
                excelMatrix.ModifyForChangeInSublines(maximumSublineCount);
            }

            var singleSegmentExcelMatrices = _segment.ExcelMatrices.OfType<SingleOccurrenceSegmentExcelMatrix>();
            foreach (var excelMatrix in singleSegmentExcelMatrices)
            {
                excelMatrix.MoveRangesWhenSublinesChange(maximumSublineCount);
            }

            var singleProfileExcelMatrices = _segment.ExcelMatrices.OfType<SingleOccurrenceProfileExcelMatrix>();
            foreach (var excelMatrix in singleProfileExcelMatrices)
            {
                excelMatrix.MoveRangesWhenSublinesChange(maximumSublineCount);
            }
        }

        private static void SetAdditionalHeaderRangeValues(string rangeName, string[,] values)
        {
            var rangeSubset = rangeName.GetRangeSubset(0, 1);
            rangeSubset.Value = values;
        }
    }
}
