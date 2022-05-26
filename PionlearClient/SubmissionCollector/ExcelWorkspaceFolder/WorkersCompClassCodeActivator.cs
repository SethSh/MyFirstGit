using System;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Profiles.ExcelComponent.Helpers;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class WorkersCompClassCodeActivator
    {
        public static void ToggleUtilization(IWorkbookLogger logger)
        {
            try
            {
                var validator = new SegmentWorksheetValidator();
                if (!validator.Validate()) return;
                
                var segment = validator.Segment;
                if (!segment.IsStructureModifiable)
                {
                    var message = segment.GetBlockModificationsMessage(BexConstants.WorkersCompClassCodeName);
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }

                if (!segment.IsWorkersComp) 
                {
                    throw new NotSupportedException($"Can only be used with a {BexConstants.WorkersCompName} {BexConstants.SegmentName}");
                }
                
                var profile = segment.WorkersCompClassCodeProfile;
                if (segment.IsWorkerCompClassCodeActive && !profile.ExcelMatrix.GetPureInputRangeWithoutBuffer().IsEmpty())
                {
                    MessageHelper.Show("Can't toggle off class codes unless it's blank", MessageType.Stop);
                    return;
                }

                if (!segment.IsWorkerCompClassCodeActive && !segment.WorkersCompStateHazardGroupProfile.ExcelMatrix.GetInputRange().IsAllZero())
                {
                    MessageHelper.Show("Can't toggle on class codes while the state by hazard group has non-zero values", MessageType.Stop);
                    return;
                }


                Toggle(segment, profile);

                segment.WorkersCompStateHazardGroupProfile.IsDirty = true;
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("WC Class Code Toggle Failed", MessageType.Stop);
            }
        }

        private static void Toggle(ISegment segment, WorkersCompClassCodeProfile profile)
        {
            segment.IsWorkerCompClassCodeActive = !segment.IsWorkerCompClassCodeActive;
            var stateHazardGroupExcelMatrix = segment.WorkersCompStateHazardGroupProfile.ExcelMatrix;
            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    if (!segment.IsWorkerCompClassCodeActive)
                    {
                        segment.ExcelComponents.Remove(profile);
                        stateHazardGroupExcelMatrix.GetInputRange().ClearContents();
                    }
                    else
                    {
                        segment.AddWorkersCompClassCodeProfile();
                    }

                    var helper = new WorkersCompClassCodeExcelMatrixHelper(segment);
                    helper.ModifyRange();

                    stateHazardGroupExcelMatrix.IsIndependent = false;
                    stateHazardGroupExcelMatrix.Reformat();

                    //needs to come after
                    if (segment.IsWorkerCompClassCodeActive)
                    {
                        stateHazardGroupExcelMatrix.WriteInClassCodeFormula();
                    }
                }
            }
        }
    }
}
