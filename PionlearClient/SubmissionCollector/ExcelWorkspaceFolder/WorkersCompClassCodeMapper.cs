using System;
using System.Linq;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class WorkersCompClassCodeMapper
    {
        public static void Map(IWorkbookLogger logger)
        {
            try
            {
                var validator = new SegmentWorksheetValidator();
                if (!validator.Validate()) return;

                var segment = validator.Segment;
                if (!segment.IsWorkersComp) 
                {
                    throw new NotSupportedException($"Can only be used with a {BexConstants.WorkersCompName} {BexConstants.SegmentName}");
                }
                
                var profile = segment.WorkersCompClassCodeProfile;
                if (!segment.IsWorkerCompClassCodeActive)
                {
                    MessageHelper.Show($"Can't map {BexConstants.WorkersCompClassCodeName.ToLower()}s because matrix isn't used", MessageType.Stop);
                    return;
                }

                if (profile.ExcelMatrix.GetInputRange().IsEmpty())
                {
                    MessageHelper.Show($"Can't map {BexConstants.WorkersCompClassCodeName.ToLower()}s because matrix is empty", MessageType.Stop);
                    return;
                }

                Map(profile);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("WC Class Code Mapping Failed", MessageType.Stop);
            }
        }

        internal static void Map(WorkersCompClassCodeProfile profile)
        {
            var hazardGroups = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.ToList();
            
            const int stateIndex = 0;
            const int classCodeIndex = 1;
            const int amountIndex = 2;

            var inputRange = profile.ExcelMatrix.GetInputRangeWithoutBuffer();
            var inputStateAbbreviations = inputRange.GetColumn(stateIndex).GetContent().ForceContentToStrings().GetColumn(0).ToList();
            var inputClassCodes = inputRange.GetColumn(classCodeIndex).GetContent().ForceContentToStrings().GetColumn(0).ToList();
            var rowCount = inputRange.Rows.Count;
            var result = new string[rowCount, 2];
            
            var uniqueInputStateAbbreviations = inputStateAbbreviations.Where(sa => sa != null).Distinct().ToList();
            var classCodeByStateDictionary = WorkersCompClassCodesAndHazardsFromBex.GetClassCodeByStateDictionary(uniqueInputStateAbbreviations);
            
            const int nameIndex = 0;
            const int descriptionIndex = 1;

            for (var row = 0; row < rowCount; row++)
            {
                var stateAbbreviation = inputStateAbbreviations[row];
                var stateClassCode = inputClassCodes[row];
                if (string.IsNullOrEmpty(stateAbbreviation) || stateClassCode == null) continue;

                var classCodeDictionary = classCodeByStateDictionary[stateAbbreviation];
                var stateClassCodeAsNumber = Convert.ToInt32(stateClassCode);
                if (!classCodeDictionary.ContainsKey(stateClassCodeAsNumber)) continue;

                var model = classCodeDictionary[stateClassCodeAsNumber];
                if (!model.HazardGroupId.HasValue) continue;

                var name = hazardGroups.Single(hg => hg.Id == model.HazardGroupId).Name;
                var extendedName = name;

                result[row, nameIndex] = extendedName;
                result[row, descriptionIndex] = model.StateDescription;
            }


            using (new ExcelEventDisabler())
            {
                inputRange.GetRangeSubset(0, amountIndex + 1).Value2 = result;
            }
        }
    }
}
