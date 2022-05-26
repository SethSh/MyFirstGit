using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;

namespace SubmissionCollector.Models.Segment
{
    internal static class RangeLocator
    {

        internal static Range FindPreviousRange(int segmentId, string componentName)
        {
            var list = BuildList(segmentId, componentName);
            return FindRangeInList(list);
        }

        private static LinkedList<IEnumerable<string>> BuildList(int segmentId, string componentName)
        {
            //this approach walks the worksheet from left to right
            //a better approach may have been to use the inter-display and intra-display
            //  to determine order

            var linkedList = new LinkedList<IEnumerable<string>>();
            var workbookRangeNames = Globals.ThisWorkbook.Names.Cast<Name>().ToList();
            var segmentPrefix = $"segment{segmentId}.";
            
            
            var pattern = $"{segmentPrefix}{ExcelConstants.SublineAllocationRangeName}" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.SublineAllocationName) return linkedList;

            
            pattern = $"{segmentPrefix}{ExcelConstants.UmbrellaAllocationRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.UmbrellaAllocationName) return linkedList;

            
            pattern = $"{segmentPrefix}{ExcelConstants.HazardProfileRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.HazardProfileName) return linkedList;

            pattern = $"{segmentPrefix}{ExcelConstants.PolicyProfileRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.PolicyProfileName) return linkedList;
            
            pattern = $"{segmentPrefix}{ExcelConstants.StateProfileRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.StateProfileName) return linkedList;

            
            pattern = $"{segmentPrefix}{ExcelConstants.ProtectionClassProfileRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.ProtectionClassProfileName) return linkedList; 
            
            pattern = $"{segmentPrefix}{ExcelConstants.ConstructionTypeProfileRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.ConstructionTypeProfileName) return linkedList;
            
            pattern = $"{segmentPrefix}{ExcelConstants.OccupancyTypeProfileRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.OccupancyTypeProfileName) return linkedList;

            pattern = $"{segmentPrefix}{ExcelConstants.TotalInsuredValueProfileRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.TotalInsuredValueProfileName) return linkedList;


            pattern = $"{segmentPrefix}{ExcelConstants.MinnesotaRetentionRangeName}.{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.MinnesotaRetentionName) return linkedList;

            pattern = $"{segmentPrefix}{ExcelConstants.WorkersCompClassCodeProfileRangeName}.{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.WorkersCompClassCodeProfileName) return linkedList; 
            
            pattern = $"{segmentPrefix}{ExcelConstants.WorkersCompStateHazardProfileRangeName}.{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.WorkersCompStateHazardGroupProfileName) return linkedList;
            
            pattern = $"{segmentPrefix}{ExcelConstants.WorkersCompStateAttachmentProfileRangeName}.{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.WorkersCompStateAttachmentProfileName) return linkedList;


            pattern = $"{segmentPrefix}{ExcelConstants.PeriodSetRangeName}" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.PeriodSetName) return linkedList;
            
            pattern = $"{segmentPrefix}{ExcelConstants.ExposureSetRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.ExposureSetName) return linkedList;
            
            pattern = $"{segmentPrefix}{ExcelConstants.AggregateLossSetRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.AggregateLossSetName) return linkedList;
            
            pattern = $"{segmentPrefix}{ExcelConstants.IndividualLossSetRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
            if (componentName == BexConstants.IndividualLossSetName) return linkedList;
            
            pattern = $"{segmentPrefix}{ExcelConstants.RateChangeSetRangeName}" + @"\d+" + $".{ExcelConstants.HeaderRangeName}";
            linkedList.AddLast(FindMatches(workbookRangeNames, pattern));
        
            return linkedList;
        }

        private static IEnumerable<string> FindMatches(List<Name> workbookRangeNames, string pattern)
        {
            var regex = new Regex(pattern + @"$");
            return RangeExtensions.GetMatchingRangeNames(regex, workbookRangeNames);
        }

        private static Range FindRangeInList(LinkedList<IEnumerable<string>> linkedList)
        {
            var currentNode = linkedList.Last;
            while (currentNode != null)
            {
                var rangeNames = currentNode.Value.ToList();
                if (rangeNames.Any())
                {
                    return FindRightMostRange(rangeNames);
                }

                currentNode = currentNode.Previous;
            }

            const string message = "Can't find range insert location";
            throw new ArgumentOutOfRangeException(message);
        }


        private static Range FindRightMostRange(IList<string> rangeNames)
        {
            if (rangeNames.Count == 1)
            {
                var rangeName = rangeNames.First();
                return rangeName.GetRange();
            }
            
            Range rightMostRange = null;
            foreach (var rangeName in rangeNames)
            {
                var range = rangeName.GetRange();
                if (rightMostRange == null)
                {
                    rightMostRange = range;
                }
                else
                {
                    if (rightMostRange.GetTopRightCell().Column < range.GetTopRightCell().Column)
                    {
                        rightMostRange = range;
                    }
                }
            }

            return rightMostRange;
        }
    }
}