using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.StateSort
{
    internal class StateSorterBasedOnCodeWithCwOnTop: BaseStateSorter
    {
        public override void Sort()
        {
            if (!HasCountrywide)
            {
                MessageHelper.Show("Countrywide option is not available for this matrix", MessageType.Stop);
                return;
            }

            SortColumn = 0;
            
            var countrywideOffset = 0; 
            const string countrywideAbbreviation = "CW";
            const string countrywideTemporaryAbbreviation = "AAA";

            var rangeContent = BodyRange.GetContent();
            for (var row = 0; row < rangeContent.GetLength(0); row ++)
            {
                if (rangeContent[row, SortColumn].ToString() != countrywideAbbreviation) continue;
                countrywideOffset = row;
                break;
            }
            BodyRange.GetTopLeftCell().Offset[countrywideOffset, SortColumn].Value2 = countrywideTemporaryAbbreviation;
            
            base.Sort();

            BodyRange.GetTopLeftCell().Value2 = countrywideAbbreviation;
        }
    }
}
