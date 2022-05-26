using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using PionlearClient.Extensions;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Models.DataComponents
{
    public abstract class MultipleOccurrenceSegmentExcelMatrix : BaseSegmentExcelMatrix, ICollection<ISubline>, IRangeMovable
    {
        [JsonProperty] private readonly List<ISubline> _sublines;

        protected MultipleOccurrenceSegmentExcelMatrix(int segmentId) : base(segmentId)
        {
            _sublines = new List<ISubline>();
        }

        protected MultipleOccurrenceSegmentExcelMatrix()
        {
            _sublines = new List<ISubline>();
        }

        public override string FullName => $"{FriendlyName} {GetLineSublineConcatenatedNames(_sublines)}";
        public virtual int ComponentId { get; set; }
        public override string RangeName => $"segment{SegmentId}.{ExcelRangeName}{ComponentId}";
        public virtual string SublinesRangeName => $"{RangeName}.{ExcelConstants.SublinesRangeName}";
        public virtual string SublinesHeaderRangeName => $"{SublinesRangeName}.{ExcelConstants.HeaderRangeName}";
        
        public virtual Range GetSublinesRange()
        {
            return SublinesRangeName.GetRange();
        }

        public virtual Range GetSublinesHeaderRange()
        {
            return SublinesHeaderRangeName.GetRange();
        }

        public virtual void ModifyForChangeInSublines(int sublineCount)
        {
            //insert/delete sub rows
            var segment = GetSegment();
            var multipleSegmentExcelMatrices = segment.ExcelMatrices.OfType<MultipleOccurrenceSegmentExcelMatrix>().ToList();
            var maximumSublineCount = multipleSegmentExcelMatrices.Max(excelMatrix => excelMatrix.Count);
            MoveRangesWhenSublinesChange(maximumSublineCount);

            //write in subline names
            var sublinesRange = SublinesRangeName.GetRange();
            var rangeContent = sublinesRange.GetContent();

            var originalRowCount = rangeContent.GetLength(0);
            var sublineNamesInRange = new List<string>();
            for (var row = 0; row < originalRowCount; row++)
            {
                var valueFromExcel = rangeContent[row, 0];
                if (valueFromExcel != null) sublineNamesInRange.Add(valueFromExcel.ToString());
            }

            if (this.Any())
            {
                var sublineNames = this.Select(x => x.ShortNameWithLob).ToList();
                sublineNames.Sort();
                if (sublineNames.IsNotEqualTo(sublineNamesInRange))
                {
                    sublinesRange.ClearContents();
                    sublinesRange.Resize[sublineNames.Count, 1].Value2 = sublineNames.ToNByOneArray();
                }
            }
            else
            {
                sublinesRange.ClearContents();
            }

            //set range properties
            sublinesRange.Resize[sublineCount, sublinesRange.Columns.Count].SetInvisibleRangeName(SublinesRangeName);
            sublinesRange = SublinesRangeName.GetRange();
            sublinesRange.Locked = true;
            sublinesRange.SetInputLabelInteriorColor();
            sublinesRange.SetBorderAroundToOrdinary();
            CenterSublines(sublinesRange);
        }

        public void UpdateSublines(IList<ISubline> changedSublines)
        {
            var originalSublines = this.ToList();
            originalSublines.Except(changedSublines).ForEach(x => Remove(x));
            changedSublines.Except(originalSublines).ForEach(Add);
        }

        [JsonIgnore] public abstract bool IsOkToMoveRight { get; }

        [JsonIgnore] public abstract bool IsOkToMoveLeft { get; }

        public abstract void MoveRight();
        public abstract void MoveLeft(int steps = 1);
        public abstract void SwitchDisplayOrderAfterMoveRight();
        public abstract void SwitchDisplayOrderAfterMoveLeft();


        public int FindLabelIndex(string label)
        {
            return GetLabels().FindIndex(s => s.Equals(label));
        }

        public int FindMaximumLabelIndicesMatchPartial(string label)
        {
            return GetLabels().FindAllIndexOf(label).Max(index => index);
        }

        public int FindMinimumLabelIndicesMatchPartial(string label)
        {
            return GetLabels().FindAllIndexOf(label).Min(index => index);
        }

        public IEnumerable<int> FindAllLabelIndicesWithPartialMatch(string label)
        {
            return GetLabels().FindAllIndexOf(label).OrderBy(index => index);
        }

        public bool DoLabelsContainAnExactMatch(string labelToMatch)
        {
            var labels = GetLabels();
            return labels.Any(label => label.Equals(labelToMatch, StringComparison.CurrentCultureIgnoreCase));
        }

        public bool DoLabelsContainAnExactMatch(IList<string> labelsToMatch)
        {
            var labels = GetLabels();
            foreach (var labelToMatch in labelsToMatch)
            {
                if (labels.Any(label => label.Equals(labelToMatch, StringComparison.CurrentCultureIgnoreCase)))
                {
                    return true;
                }
            }

            return false;

        }
        public IEnumerable<int> FindAllLabelIndicesMatchPartial(IList<string> labels)
        {
            var allLabels = GetLabels();

            var matches = new List<int>();
            foreach (var label in labels)
            {
                matches.AddRange(allLabels.FindAllIndexOf(label));
            }

            return matches.Distinct().OrderBy(index => index);
        }

        public static string GetRangeName(int segmentId, int componentId, string excelName)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{excelName}{componentId}";
        }

        public static string GetBasisRangeName(int segmentId, int componentId, string excelName)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{excelName}{componentId}.{ExcelConstants.ProfileBasisRangeName}";
        }

        public static string GetHeaderRangeName(int segmentId, int componentId, string excelName)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{excelName}{componentId}.{ExcelConstants.HeaderRangeName}";
        }

        public static string GetSublinesRangeName(int segmentId, int componentId, string excelName)
        {
            return $"{ExcelConstants.SegmentRangeName}{segmentId}.{excelName}{componentId}.{ExcelConstants.SublinesRangeName}";
        }

        public static string GetSublinesHeaderRangeName(int segmentId, int componentId, string excelName)
        {
            return $"{GetSublinesRangeName(segmentId, componentId, excelName)}.{ExcelConstants.HeaderRangeName}";
        }

        private List<string> GetLabels()
        {
            var labelsRange = GetInputLabelRange();
            return labelsRange.GetContent().ForceContentToStrings().GetRow(0).ToList();
        }

        private static void CenterSublines(Range sublinesRange)
        {
            if (sublinesRange.Columns.Count == 1)
            {
                sublinesRange.AlignCenter();
                return;
            }

            sublinesRange.AlignCenterAcrossSelection();
        }

        #region list management

        public IEnumerator<ISubline> GetEnumerator()
        {
            return _sublines.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(ISubline item)
        {
            _sublines.Add(item);
        }

        public void Clear()
        {
            _sublines.Clear();
        }

        public bool Contains(ISubline item)
        {
            var match = _sublines.FirstOrDefault(sl => item != null && sl.Code == item.Code);
            return match != null;
            //the rebuild caused a problem because we're added a clone ... and the contains was failing
            //Contains(item);
        }

        public void CopyTo(ISubline[] array, int arrayIndex)
        {
            _sublines.CopyTo(array, arrayIndex);
        }

        public bool Remove(ISubline item)
        {
            return _sublines.Remove(item);
        }

        public int Count => _sublines.Count;
        public bool IsReadOnly => false;

        #endregion
    }
}