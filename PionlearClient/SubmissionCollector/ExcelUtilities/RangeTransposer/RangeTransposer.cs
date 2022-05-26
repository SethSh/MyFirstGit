using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.ExcelUtilities.RangeTransposer
{
    public class RangeTransposer: BaseRangeTransposer
    {
        public RangeTransposer(ISegment segment, IExcelMatrix excelMatrix) : base(segment, excelMatrix)
        {
                
        }

        public override void Transpose()
        {
            //insert columns to right
            //insert cells below  - this will be your scratchpad, otherwise you could cut over existing cells
            //move labels + values
            //move row header
            //cut up
            //delete scratchpad cells
            //rename ranges and reformat
                
            var inputRange = ExcelMatrix.GetInputRange();
            
            var rowLabels = ExcelMatrix.GetInputLabelRange().GetContent();
            var columnLabels = BodyHeaderRange.GetContent();
            
            var rowCount = rowLabels.GetLength(0);
            var columnCount = columnLabels.GetLength(1);

            BodyRange.Locked = false;
            
            inputRange.GetTopRightCell().Offset[0, 1].Resize[1, rowCount-columnCount + 1].InsertColumnsToRight();
            BodyRange.GetBottomLeftCell().Offset[1, 0].Resize[columnCount, 1].InsertRangeDown();
            
            ((IRangeTransposable)ExcelMatrix).IsTransposed = !((IRangeTransposable)ExcelMatrix).IsTransposed;
            using (new ExcelEventDisabler())
            {

                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        var s = BodyRange.GetTopLeftCell();
                        var d = BodyRange.GetBottomRightCell();

                        s.Offset[row, column].Cut(d.Offset[column +1, row - (columnCount-2)]);
                    }
                }

                for (var column = 0; column < columnCount; column++)
                {
                    var s = BodyHeaderRange.GetTopLeftCell();
                    var d = BodyRange.GetBottomLeftCell();
                    s.Offset[0, column].Cut(d.Offset[column + 1, 0]);
                }
                
                HeaderRange.Offset[0, columnCount - 1].ClearContents();
                HeaderRange.Offset[1,0].Resize[rowCount+1, rowCount+1].DeleteRangeUp();
                HeaderRange.Resize[1, rowCount + 1].SetInvisibleRangeName(((ISegmentExcelMatrix)ExcelMatrix).HeaderRangeName);
                HeaderRange.Offset[1, 0].Resize[columnCount, rowCount + 1].SetInvisibleRangeName(ExcelMatrix.RangeName);

                if (ExcelMatrix is MultipleOccurrenceSegmentExcelMatrix excelComponentWithSublines 
                    && !(excelComponentWithSublines is HazardExcelMatrix)
                    && !(excelComponentWithSublines is ConstructionTypeExcelMatrix)
                    && !(excelComponentWithSublines is ProtectionClassExcelMatrix))
                {
                    var anchor = excelComponentWithSublines.GetSublinesHeaderRange().GetTopLeftCell();
                    anchor.Resize[1, rowCount + 1].SetInvisibleRangeName(excelComponentWithSublines.SublinesHeaderRangeName);

                    anchor = excelComponentWithSublines.GetSublinesRange().GetTopLeftCell();
                    anchor.Resize[Segment.Count, rowCount + 1].SetInvisibleRangeName(excelComponentWithSublines.SublinesRangeName);
                }

                ExcelMatrix.Reformat();
            }
        }
    }
}
