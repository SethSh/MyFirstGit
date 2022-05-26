using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;
using SubmissionCollector.View.Enums;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class UmbrellaTypeAllocatorManager
    {
        internal FormResponse SelectUmbrellaTypes(IWorkbookLogger logger)
        {
            try
            {
                return SelectUmbrellaTypes();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Select {BexConstants.UmbrellaTypeName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
                return FormResponse.Cancel;
            }
        }

        internal FormResponse SelectUmbrellaTypes()
        {
            var rangeValidator = new SegmentWorksheetValidator();
            if (!rangeValidator.Validate()) return FormResponse.Cancel;

            var segment = rangeValidator.Segment;
            if (!segment.IsStructureModifiable)
            {
                var message = segment.GetBlockModificationsMessage(BexConstants.UmbrellaTypeName);
                MessageHelper.Show(message, MessageType.Stop);
                return FormResponse.Cancel;
            }

            if (!segment.AreUmbrellaTypesModifiable)
            {
                var message = $"In order to modify the {BexConstants.UmbrellaTypeName.ToLower()}s list, " +
                              $"all commercial {BexConstants.PolicyProfileName.ToLower()}s must treat {BexConstants.UmbrellaTypeName.ToLower()} as grouped.";
                MessageHelper.Show(message, MessageType.Stop);
                return FormResponse.Cancel;
            }

            var text = $"{BexConstants.SegmentName} Selector";
            var viewModel = new UmbrellaTypeAllocatorViewModel(segment);
            var allocator = new UmbrellaTypeAllocator(viewModel);
            var form = new UmbrellaTypeAllocationForm(allocator)
            {
                Text = text,
                Height = (int) (((int) FormSizeHeight.Medium + (int) FormSizeHeight.Small) / 2d),
                Width = (int) ((int) FormSizeWidth.Small * 0.75),
                StartPosition = FormStartPosition.CenterScreen
            };
            form.ShowDialog();

            if (allocator.Response != FormResponse.Ok) return FormResponse.Cancel;

            var selectedCodes = viewModel.UmbrellaItems.Where(item => item.IsSelected).Select(item => item.UmbrellaTypeCode.ToString());
            ModifyWorksheetFromUmbrellaChanges(segment, selectedCodes.ToList());
            return FormResponse.Ok;
        }

        internal static void ModifyWorksheetFromUmbrellaChanges(ISegment segment, IList<string> selectedCodes)
        {
            using (new ExcelScreenUpdateDisabler())
            {
                using (new ExcelEventDisabler())
                {
                    using (new WorkbookUnprotector())
                    {
                        segment.WorksheetManager.CreateUmbrellaMatrix();

                        var umbrellaAllocationInRange = segment.WorksheetManager.GetUmbrellaMatrixContent();
                        var umbrellaNamesFromRange = umbrellaAllocationInRange.Keys.ToList();
                        var umbrellaCodesFromRange = UmbrellaTypesFromBex.GetCodes(umbrellaNamesFromRange)
                            .Select(item => item.ToString())
                            .ToList();

                        segment.WorksheetManager.WriteUmbrellaNamesIntoExistingRange(selectedCodes, umbrellaCodesFromRange);
                    }
                }
            }
        }
    }
}