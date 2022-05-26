using System.Linq;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.ExcelUtilities.Extensions;

namespace SubmissionCollector.Models.DataComponents
{
    public abstract class MultipleOccurrenceSegmentLossExcelMatrix : MultipleOccurrenceSegmentExcelMatrix
    {
        protected MultipleOccurrenceSegmentLossExcelMatrix(int segmentId) : base(segmentId)
        {
                
        }

        protected MultipleOccurrenceSegmentLossExcelMatrix()
        {
                
        }

        public void ModifyRangeToReflectChangeToPaid(bool isPaidAvailable)
        {
            if (isPaidAvailable)
            {
                InsertPaidColumns();
            }
            else
            {
                DeletePaidColumns();
            }
        }

        

        protected abstract void DeletePaidColumns();

        protected abstract void InsertPaidColumns();

        public void ModifyRangeToReflectChangeToAlaeFormat(bool isLossAndAlaeCombined)
        {
            if (isLossAndAlaeCombined)
            {
                DeleteAlaeColumns();
            }
            else
            {
                InsertAlaeColumns();
            }
        }

        protected void InsertAlaeColumns()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.ReportedLossName)) return; 
            
            var labelRange = GetInputLabelRange();
            var columnIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.LossAndAlaeName);
            columnIndices.Reverse().ForEach(index =>
            {
                labelRange.Offset[0, index + 1].Resize[1, 1].InsertColumnsToRight();

                labelRange = GetInputLabelRange();

                var labelRangeInLoop = labelRange.GetTopLeftCell().Offset[0, index];

                var combinedLabel = labelRangeInLoop.Value2.ToString();
                var lossLabel = combinedLabel.Replace(BexConstants.LossAndAlaeName, BexConstants.LossName);
                var alaeLabel = combinedLabel.Replace(BexConstants.LossAndAlaeName, BexConstants.AlaeName);

                labelRangeInLoop.Offset[0, 0].Value2 = lossLabel;
                labelRangeInLoop.Offset[0, 1].Value2 = alaeLabel;
            });

            GetSublinesHeaderRange().AppendColumn().SetInvisibleRangeName(SublinesHeaderRangeName);
            GetSublinesRange().AppendColumn().SetInvisibleRangeName(SublinesRangeName);

            GetHeaderRange().AppendColumn().SetInvisibleRangeName(HeaderRangeName);
            RangeName.GetRange().AppendColumn().SetInvisibleRangeName(RangeName);
        }

        protected void DeleteAlaeColumns()
        {
            if (DoLabelsContainAnExactMatch(BexConstants.ReportedLossAndAlaeName)) return; 
            
            var labelsRange = GetInputLabelRange();
            var columnIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.AlaeName);
            columnIndices.Reverse().ForEach(index => labelsRange.Resize[1, 1].Offset[0, index].EntireColumn.Delete());

            var labelsRangePostDelete = GetInputLabelRange();
            columnIndices = FindAllLabelIndicesWithPartialMatch(BexConstants.LossName);
            columnIndices.ForEach(index =>
            {
                var labelRange = labelsRangePostDelete.Resize[1, 1].Offset[0, index];
                var label = labelRange.Value2.ToString();
                labelRange.Value2 = label.Replace(BexConstants.LossName, BexConstants.LossAndAlaeName);
            });
        }
    }
}
