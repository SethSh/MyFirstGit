using System;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class WorkersCompStateByHazardGroupIndependentActivator
    {
        public static void ToggleIndependence(IWorkbookLogger logger)
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

                if (segment.IsWorkerCompClassCodeActive)
                {
                    MessageHelper.Show("Can't toggle state by hazard group independence when class codes are utilized", MessageType.Stop);
                    return;
                }

                var profile = segment.WorkersCompStateHazardGroupProfile;
                var excelMatrix = profile.ExcelMatrix;
                var inputRange = excelMatrix.GetInputRange();
                
                var failNonZeroValidation1 = !excelMatrix.IsIndependent && inputRange.IsNotAllZero();
                var failNonZeroValidation2 = excelMatrix.IsIndependent
                                             && (excelMatrix.GetStatePremiumsRange().IsNotAllZero()
                                             || excelMatrix.GetHazardGroupRange().Offset[-1, 0].IsNotAllZero());
                
                if (failNonZeroValidation1 || failNonZeroValidation2)
                {
                    MessageHelper.Show("Can't toggle state by hazard group independence while the input range has non-zero values", MessageType.Stop);
                    return;
                }

                Toggle(profile);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("WC State By Hazard Group Independence Toggle Failed", MessageType.Stop);
            }
        }

        private static void Toggle(WorkersCompStateHazardGroupProfile profile)
        {
            var excelMatrix = profile.ExcelMatrix;
            excelMatrix.IsIndependent = !excelMatrix.IsIndependent;

            var inputRange = excelMatrix.GetInputRange();
            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    if (excelMatrix.IsIndependent)
                    {
                        excelMatrix.GetHazardGroupPercentRange().ClearContents();
                        excelMatrix.GetStatePremiumsRange().ClearContents();
                    }
                    else
                    {
                        inputRange.ClearContents();
                    }

                    excelMatrix.Reformat();
                }
            }

            profile.IsDirty = true;
        }
    }
}
