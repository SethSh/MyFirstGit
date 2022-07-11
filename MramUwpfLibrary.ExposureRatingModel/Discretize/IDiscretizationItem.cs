using System;

namespace MramUwpfLibrary.ExposureRatingModel.Discretize
{
    public class DiscretizationItem : IDiscretizationItem
    {
        public double Loss { get; set; }
        public double Probability { get; set; }
        public object Clone()
        {
            return new DiscretizationItem { Loss = Loss, Probability = Probability };
        }

        public override string ToString()
        {
            return $"Loss: {Loss:N0}    Prob: {Probability:N4}";
        }
    }

    public interface IDiscretizationItem : ICloneable
    {
        double Loss { get; set; }
        double Probability { get; set; }
    }
}