using System;
using System.Linq;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Profiles.ExcelComponent.Helpers;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class UmbrellaTypePolicyProfileWizardManager
    {
        internal static void ModifyUmbrellaType(IWorkbookLogger logger)
        {
            try
            {
                ModifyUmbrellaType();
            }

            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Modify {BexConstants.UmbrellaTypeName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void ModifyUmbrellaType()
        {
            var rangeValidator = new SegmentWorksheetValidator();
            if (!rangeValidator.Validate()) return;

            var segment = rangeValidator.Segment;
            if (!segment.IsStructureModifiable)
            {
                var message = segment.GetBlockModificationsMessage(BexConstants.UmbrellaTypeName);
                MessageHelper.Show(message, MessageType.Stop);
                return;
            }

            var umbrellaValidator = new UmbrellaValidator();
            if (!umbrellaValidator.Validate(segment)) return;

            var viewModel = new UmbrellaTypePolicyProfileWizardViewModel(segment);
            var wizard = new UmbrellaTypePolicyProfileWizard(viewModel);
            var form = new UmbrellaTypePolicyProfileWizardForm(wizard)
            {
                Text = @"Submission Segment Umbrella Type Wizard",
                Height = (int) FormSizeHeight.ExtraLarge,
                Width = (int) FormSizeWidth.ExtraLarge,
                StartPosition = FormStartPosition.CenterScreen
            };
            form.ShowDialog();

            if (wizard.Response != FormResponse.Ok) return;

            MapViewModel(segment, viewModel);
            ModifyWorksheet(segment);
        }

        private static void MapViewModel(ISegment segment, IUmbrellaPolicyProfileWizardViewModel viewModel)
        {
            #region add/remove policy profiles

            foreach (var viewModelPolicyProfile in viewModel.ComponentViews.Where(x => x.IsVisible))
            {
                var policyGuid = viewModelPolicyProfile.Guid;
                var policyComponentId = viewModelPolicyProfile.Id;
                var displayOrder = viewModelPolicyProfile.DisplayOrder;

                int? umbrellaTypeCode = null;
                if (viewModelPolicyProfile.UmbrellaTypes.Count > 1)
                {
                    umbrellaTypeCode = Convert.ToInt16(viewModelPolicyProfile.UmbrellaTypes.First().UmbrellaTypeCode);
                }

                if (segment.PolicyProfiles.SingleOrDefault(x => x.Guid == policyGuid) == null)
                {
                    var policyProfile = new PolicyProfile(segment.Id, policyComponentId)
                    {
                        Name = viewModelPolicyProfile.FullName,
                        Guid = policyGuid,
                        UmbrellaType = umbrellaTypeCode
                    };
                    policyProfile.ExcelMatrix.IntraDisplayOrder = viewModelPolicyProfile.DisplayOrder;
                    segment.ExcelComponents.Add(policyProfile);
                }
                else
                {
                    var existingPolicyProfile = segment.PolicyProfiles.SingleOrDefault(y => y.Guid == policyGuid);
                    if (existingPolicyProfile == null) continue;

                    var policyExcelMatrix = existingPolicyProfile.ExcelMatrix;
                    if (policyExcelMatrix.IntraDisplayOrder != displayOrder)
                    {
                        policyExcelMatrix.IntraDisplayOrder = displayOrder;
                    }
                }
            }

            var viewModelPolicyProfileGuids = viewModel.ComponentViews.Where(x => x.IsVisible).Select(x => x.Guid).ToList();
            segment.ExcelComponents.RemoveAll(ec => ec is PolicyProfile && !viewModelPolicyProfileGuids.Contains(ec.Guid));

            #endregion

            var segmentUmbrellaTypes = viewModel.SegmentUmbrellaTypes;
            segment.PolicyProfiles.ForEach(policyProfile =>
            {
                var wizardComponent = viewModel.ComponentViews.Single(x => x.Guid == policyProfile.Guid);
                var wizardComponentSublines = wizardComponent.Sublines;

                if (!policyProfile.Name.Equals(wizardComponent.FullName))
                {
                    policyProfile.Name = wizardComponent.FullName;
                }

                if (policyProfile.ExcelMatrix.IsNotEqualTo(wizardComponentSublines))
                {
                    policyProfile.ExcelMatrix.UpdateSublines(wizardComponentSublines);
                }

                var umbrellaTypes = wizardComponent.UmbrellaTypes;
                policyProfile.UmbrellaType = umbrellaTypes.IsEqualTo(segmentUmbrellaTypes)
                    ? new int?()
                    : Convert.ToInt16(umbrellaTypes.First().UmbrellaTypeCode);
            });
        }

        private static void ModifyWorksheet(ISegment segment)
        {
            var helper = new PolicyExcelMatrixHelper(segment);

            using (new ExcelScreenUpdateDisabler())
            {
                using (new ExcelEventDisabler())
                {
                    helper.ModifyRanges();
                    foreach (var policyProfile in segment.PolicyProfiles)
                    {
                        var excelMatrix = policyProfile.ExcelMatrix;
                        excelMatrix.HeaderRangeName.GetTopLeftCell().Value2 = policyProfile.Name;
                        excelMatrix.SublinesHeaderRangeName.GetTopLeftCell().Value2 = $"{policyProfile.Name} Sublines";
                        excelMatrix.ModifyForChangeInSublines(segment.Count);
                    }

                    segment.UmbrellaExcelMatrix.Reformat();
                }
            }
        }
    }
}