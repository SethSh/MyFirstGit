using System.Collections.Generic;

namespace MramUwpfLibrary.ExposureRatingModel.Property
{
    internal interface ITotalInsuredValueItem
    {
        double TotalInsuredValue { get; set; }
        double? Share { get; set; }
        double? Limit { get; set; }
        double? Attachment { get; set; }
        double Weight { get; set; }
        IList<IPropertyCurveProcessor> PropertyCurveProcessors { get; set; }
    }

    internal class TotalInsuredValueItem : ITotalInsuredValueItem
    {
        public TotalInsuredValueItem()
        {
            PropertyCurveProcessors = new List<IPropertyCurveProcessor>();
        }

        public double? Share { get; set; }
        public double TotalInsuredValue { get; set; }
        public double? Limit { get; set; }
        public double? Attachment { get; set; }
        public double Weight { get; set; }
        public IList<IPropertyCurveProcessor> PropertyCurveProcessors { get; set; }
    }
}