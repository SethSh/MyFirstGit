using System.Linq;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ExcelUtilities
{
    internal class ExcelSheetActivateEventManager
    {
        internal void MonitorSheetChange(Worksheet worksheet, Package package)
        {
            if (worksheet == null)
            {
                HideSegmentRibbonItems();
                return;
            }

            if (package.Worksheet.Name == worksheet.Name)
            {
                package.IsSelected = true;
                HideSegmentRibbonItems();
                return;
            }

            var segment = package.Segments.SingleOrDefault(p => p.WorksheetManager.Worksheet.Name == worksheet.Name);
            if (segment == null)
            {
                HideSegmentRibbonItems();
                return;
            }

            segment.IsSelected = true;
            RefreshRibbon(segment);
        }

        public static void RefreshRibbon(ISegment segment)
        {
            RefreshUmbrellaWizardButton(segment);
            RefreshStateSortOptions(segment);
            RefreshLineOfBusinessMenus(segment);
        }

        public static void RefreshUmbrellaWizardButton(ISegment segment)
        {
            var show = segment.IsUmbrella && segment.ContainsAnyCommercialSublines;
            ShowUmbrellaWizardButton(show);
        }

        public static void RefreshLineOfBusinessMenus(ISegment segment)
        {
            ShowLiabilityMenu(segment.IsLiability);
            ShowPropertyMenu(segment.IsProperty);
            ShowWorkersCompMenu(segment.IsWorkersComp);
        }

        public static void RefreshStateSortOptions(ISegment segment)
        {
            var showCountrywide = !segment.IsWorkersComp;
            ShowCountrywideOptions(showCountrywide);
        }

        private static void HideSegmentRibbonItems()
        {
            ShowUmbrellaWizardButton(false);
            ShowCountrywideOptions(false);
            HideLineOfBusinessMenus();
        }

        private static void HideLineOfBusinessMenus()
        {
            ShowLiabilityMenu(false);
            ShowPropertyMenu(false);
            ShowWorkersCompMenu(false);
        }
        
        private static void ShowLiabilityMenu(bool show)
        {
            Globals.Ribbons.SubmissionRibbon.PolicyProfileDimensionMenu.Enabled = show;
        }

        private static void ShowPropertyMenu(bool show)
        {
            Globals.Ribbons.SubmissionRibbon.ChangeTivRange.Enabled = show;
        }

        private static void ShowWorkersCompMenu(bool show)
        {
            Globals.Ribbons.SubmissionRibbon.WorkercCompMenu.Visible = show;
        }

        private static void ShowUmbrellaWizardButton(bool show)
        {
            Globals.Ribbons.SubmissionRibbon.UmbrellaSplitButton.Visible = show;
        }

        private static void ShowCountrywideOptions(bool show)
        {
            var screenTip = show ? string.Empty : "Countywide options don't apply";

            Globals.Ribbons.SubmissionRibbon.SortByStateIdWithCwOnTopButton.Enabled = show;
            Globals.Ribbons.SubmissionRibbon.SortByStateIdWithCwOnTopButton.ScreenTip = screenTip;
            Globals.Ribbons.SubmissionRibbon.SortByStateNameWithCwOnTopButton.Enabled = show;
            Globals.Ribbons.SubmissionRibbon.SortByStateIdWithCwOnTopButton.ScreenTip = screenTip;
        }
    }
}