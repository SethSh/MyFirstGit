using System.Linq;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.StateSort
{
    internal abstract class BaseStateSorter : IStateSorter
    {
        public int SortColumn { get; set; }
        public ISegment Segment { get; set; }
        public ISegmentExcelMatrix ExcelMatrix { get; set; }
        public string SelectedRangeName { get; set; }
        public Range BodyRange { get; set; }
        public bool HasCountrywide { get; set; }

        public bool Validate()
        {
            var rangeValidator = new SegmentWorksheetValidator();
            if (!rangeValidator.Validate()) return false;

            var selectedRange = rangeValidator.SelectedRange;
            Segment = rangeValidator.Segment;

            var profiles = Segment.ExcelComponents.Where(ec => ec is IProvidesState).ToList();
            var rangeNames = profiles.Select(x => x.CommonExcelMatrix.RangeName);
            SelectedRangeName = string.Empty;
            foreach (var item in rangeNames)
            {
                if (!item.ContainsRange(selectedRange)) continue;
                SelectedRangeName = item;
                break;
            }

            if (string.IsNullOrEmpty(SelectedRangeName))
            {
                const string message = "The selection must be within a profile containing state";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            var profile = profiles.Single(x => x.CommonExcelMatrix.RangeName == SelectedRangeName);
            HasCountrywide = profile is StateProfile;
            ExcelMatrix = profile.CommonExcelMatrix;
            BodyRange = ExcelMatrix.GetBodyRange();
            
            return true;
        }

        public virtual void Sort()
        {
            Segment.WorksheetManager.Worksheet.Sort.SortFields.Clear();
            Segment.WorksheetManager.Worksheet.Sort.SortFields.Add(Key: BodyRange.GetColumn(SortColumn));
            Segment.WorksheetManager.Worksheet.Sort.SetRange(BodyRange);
            Segment.WorksheetManager.Worksheet.Sort.Apply();
        }

        
    }
}
