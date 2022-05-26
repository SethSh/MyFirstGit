using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.ExcelUtilities.RangeTransposer
{
    public class RangeReverseTransposer : BaseRangeTransposer
    {
        public RangeReverseTransposer(ISegment segment, IExcelMatrix excelMatrix) : base (segment, excelMatrix)
        {
                
        }

        public override void Transpose()
        {
            //insert cells below  - this will be your scratchpad, otherwise you could cut over existing cells
            //move labels + values
            //move column header
            //delete old empty space up
            //delete columns
            //rename ranges and reformat
            
            var rowLabels = ((ISegmentExcelMatrix)ExcelMatrix).GetBodyHeaderRange().GetContent();
            var columnLabels = ExcelMatrix.GetInputLabelRange().GetContent();

            var rowCount = rowLabels.GetLength(0);
            var columnCount = columnLabels.GetLength(1);

            BodyRange.Locked = false;

            BodyRange.Offset[rowCount, 0].Resize[columnCount + 1, columnCount + 1].InsertRangeDown();

            ((IRangeTransposable)ExcelMatrix).IsTransposed = !((IRangeTransposable)ExcelMatrix).IsTransposed;
            using (new ExcelEventDisabler())
            {
                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        var s = BodyRange.GetTopLeftCell();
                        var d = BodyHeaderRange.GetBottomLeftCell();

                        s.Offset[row, column].Cut(d.Offset[column + 2, row]);
                    }
                }

                for (var row = 0; row < rowCount; row++)
                {
                    var s = BodyHeaderRange.GetTopLeftCell();
                    var d = BodyHeaderRange.GetBottomLeftCell();
                    s.Offset[row, 0].Cut(d.Offset[1, row]);
                }

                HeaderRange.Offset[1,0].Resize[rowCount,columnCount+1].DeleteRangeUp();
                
                HeaderRange.GetTopLeftCell().Offset[0, rowCount].Resize[1, columnCount - rowCount +1].EntireColumn.Delete();
                HeaderRange.Offset[1, 0].Resize[columnCount + 1, rowCount].SetInvisibleRangeName(ExcelMatrix.RangeName);
                
                ExcelMatrix.Reformat();
            }
        }
    }
}