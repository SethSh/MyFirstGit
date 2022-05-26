using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal interface IExcelMatrixHelper<T> where T : IExcelComponent
    {
        ISegment Segment { set; }
        IEnumerable<T> ExcelComponents { get; }
    }

    internal abstract class BaseExcelMatrixHelper<T> : IExcelMatrixHelper<T> where T : IExcelComponent
    {
        protected BaseExcelMatrixHelper(ISegment segment)
        {
            Segment = segment;
            UserPrefs = UserPreferences.ReadFromFile();
            IsPreviousRangeAdjacent = false;
        }

        public UserPreferences UserPrefs { get; set; }
        public ISegment Segment { get; set; }
        public IEnumerable<T> ExcelComponents => Segment.ExcelComponents.OfType<T>();
        public bool IsPreviousRangeAdjacent { get; set; }
        public string ExcelRangeName { get; set; }
        public string ComponentName { get; set; }
        public int ColumnCount { get; set; }
        public int ColumnCountPlusOne => ColumnCount + 1;


        protected Range GetAdjacentRange()
        {
            var previousRange = RangeLocator.FindPreviousRange(Segment.Id, ComponentName);
            return IsPreviousRangeAdjacent ? previousRange.GetTopRightCell() : previousRange.GetTopRightCell().Offset[0, 1];
        }
    }
}
