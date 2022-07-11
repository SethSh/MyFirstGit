using System;
using System.Collections.Generic;
using System.Linq;

namespace MramUwpfLibrary.ExposureRatingModel.Property
{
    public interface IPropertyCurveProvider
    {
        //reasonable values
            //totalInsuredValue 1,000,000
            //modification factor 1.25
            //mixed exponential (property curve parameters)
                //means[]   0, 1000, 10000, 100000, 100000
                //weights[] 0.5, 0.3, 0.15, 0.04, 0.01

        IPropertyCurve GetCommercialPropertyCurve(DateTime date, double totalInsuredValue, long occupancyTypeId);
        IPropertyCurve GetPersonalPropertyCurve(DateTime date, double totalInsuredValue, long constructionTypeId, 
            long protectionClassId);
    }

    public interface IPropertyCurveFactorProvider
    {
        //reasonable value is 1.25
        double? GetPropertyFactor(DateTime date, long protectionClassId, long constructionTypeId);
    }

    public interface IPropertyAlaeToLossProvider
    {
        //reasonable value is 3.5%
        double? GetPropertyAlaeToLossFactor(DateTime date, long sublineId);
    }


    
    public interface IPropertyCurve: ICloneable
    {
        PropertyCurveParameters MixedExponentialCurveParameters { get; set; }
    }

    public class PropertyCurve: IPropertyCurve
    {
        public PropertyCurve()
        {
            MixedExponentialCurveParameters = new PropertyCurveParameters();
        }

        public PropertyCurveParameters MixedExponentialCurveParameters { get; set; }
        public object Clone()
        {
            var clone = new PropertyCurve
            {
                MixedExponentialCurveParameters =
                {
                    Means = this.MixedExponentialCurveParameters.Means.Select(d => d).ToArray(),
                    Weights = this.MixedExponentialCurveParameters.Weights.Select(d => d).ToArray()
                }
            };
            
            return clone;
        }
    }

    internal interface IPropertyCurveProcessor 
    {
        double Weight { get; set; }
        IPropertyCurve PropertyCurve { get; set; }
    }

    internal class PropertyCurveProcessor : IPropertyCurveProcessor
    {

        public PropertyCurveProcessor(IPropertyCurve propertyCurve)
        {
            PropertyCurve = propertyCurve;
        }
        public double Weight { get; set; }
        public IPropertyCurve PropertyCurve { get; set; }
    }


    public class PropertyCurveSet
    {
        public string SegmentId { get; set; }
        public string SublineId { get; set; }
        internal IList<ITotalInsuredValueItem> TotalInsuredValueItems { get; set; }
    }
}