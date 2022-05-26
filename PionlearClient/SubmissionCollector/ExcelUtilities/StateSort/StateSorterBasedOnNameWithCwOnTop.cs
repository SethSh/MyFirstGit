using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.StateSort
{
    internal class StateSorterBasedOnNameWithCwOnTop : BaseStateSorter
    {
        public override void Sort()
        {
            if (!HasCountrywide)
            {
                MessageHelper.Show("Countrywide option is not available for this matrix", MessageType.Stop);
                return;
            }
            
            SortColumn = 1;

            var countrywideOffset = 0;
            const string countrywide = "Countrywide";
            const string temporaryCountrywide = "AAA";

            var rangeContent = BodyRange.GetContent();
            for (var row = 0; row < rangeContent.GetLength(0); row++)
            {
                if (rangeContent[row, SortColumn].ToString() != countrywide) continue;
                countrywideOffset = row;
                break;
            }
            BodyRange.Resize[1, 1].Offset[countrywideOffset, SortColumn].Value2 = temporaryCountrywide;

            base.Sort();
            
            BodyRange.GetTopLeftCell().Offset[0, SortColumn].Value2 = countrywide;
        }
    }
}
