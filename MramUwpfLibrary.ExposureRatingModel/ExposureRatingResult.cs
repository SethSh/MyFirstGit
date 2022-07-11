using System.Collections.Generic;
using System.Text;

namespace MramUwpfLibrary.ExposureRatingModel
{
    public interface IExposureRatingResult
    {
        string LayerId { get; set; }
        string SubmissionSegmentId { get; set; }
        IEnumerable<IExposureRatingResultItem> Items { get; set; }
    }
    
    public class ExposureRatingResult : IExposureRatingResult
    {
        //this is technically the id that association submission segment with analysis segment
        public string LayerId { get; set; }
        public string SubmissionSegmentId { get; set; }

        public IEnumerable<IExposureRatingResultItem> Items { get; set; }

        public new string ToString()
        {
            var sb = new StringBuilder();
            foreach (var result in Items)
            {
                var s = ((ExposureRatingResultItem)result).ToString();
                sb.AppendLine(s);
            }
            return sb.ToString();
        }
    }
}
