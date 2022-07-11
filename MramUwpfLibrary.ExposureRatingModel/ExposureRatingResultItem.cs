using System;
using System.Collections.Generic;
using System.Text;

namespace MramUwpfLibrary.ExposureRatingModel
{
    public interface IExposureRatingResultItem
    {
        string SublineId { get; set; }
        double LayerLossCostPercent { get; set; }
        double AlaeToLoss { get; set; }
        double BenchmarkAlaeToLoss { get; set; }
        double LayerLossCostAmount { get; set; }
        double Severity { get; set; }
        double Frequency { get; set; }
        double UnlimitedLossPlusAlaeRatio { get; set; }
        double BenchmarkUnlimitedLossPlusAlaeRatio { get; set; }
    }

    public class ExposureRatingResultItem : IExposureRatingResultItem
    {
        public string SublineId { get; set; }
        public double LayerLossCostPercent { get; set; }
        public double AlaeToLoss { get; set; }
        public double BenchmarkAlaeToLoss { get; set; }
        public double LayerLossCostAmount { get; set; }
        public double Severity { get; set; }
        public double Frequency { get; set; }
        public double UnlimitedLossPlusAlaeRatio { get; set; }
        public double BenchmarkUnlimitedLossPlusAlaeRatio { get; set; }
        public double Weight { get; set; }

        public void CheckForNan()
        {
            var nanList = new List<string>();
            if (double.IsNaN(LayerLossCostPercent)) nanList.Add("Layer Loss Cost Percent");
            if (double.IsNaN(AlaeToLoss)) nanList.Add("Alae To Loss Ratio");
            if (double.IsNaN(BenchmarkAlaeToLoss)) nanList.Add("Benchmark Alae To Loss Ratio");
            if (double.IsNaN(LayerLossCostAmount)) nanList.Add("Layer Loss Cost Amount");
            if (double.IsNaN(Severity)) nanList.Add("Average Severity");
            if (double.IsNaN(Frequency)) nanList.Add("Frequency");
            if (double.IsNaN(UnlimitedLossPlusAlaeRatio)) nanList.Add("Unlimited Loss And Alae Ratio");
            if (double.IsNaN(BenchmarkUnlimitedLossPlusAlaeRatio)) nanList.Add("Unlimited Loss And Alae Ratio With Unadjusted Alae");

            if (nanList.Count > 0) throw new ArgumentException("Exposure Rating Results failed for " + string.Join(", ", nanList));
        }

        public new string ToString()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Subline of Business ID: " + SublineId);
            sb.AppendLine("LayerLossCostPercent: " + LayerLossCostPercent.ToString("N10"));
            sb.AppendLine("AlaeToLoss: " + AlaeToLoss.ToString("N10"));
            sb.AppendLine("BenchmarkAlaeToLoss: " + BenchmarkAlaeToLoss.ToString("N10"));
            sb.AppendLine("LayerLossCostAmount: " + LayerLossCostAmount.ToString("N10"));
            sb.AppendLine("AverageSeverity: " + Severity.ToString("N10"));
            sb.AppendLine("Frequency: " + Frequency.ToString("N10"));
            sb.AppendLine("UnlimitedLossPlusAlaeRatio: " + UnlimitedLossPlusAlaeRatio.ToString("N10"));
            sb.AppendLine("BenchmarkUnlimitedLossPlusAlaeRatio: " + BenchmarkUnlimitedLossPlusAlaeRatio.ToString("N10"));

            return sb.ToString();
        }

        public static IExposureRatingResultItem CopyAndAdjust(string sublineId, double factor, IExposureRatingResultItem item)
        {
            return new ExposureRatingResultItem
            {
                SublineId = sublineId,
                AlaeToLoss = item.AlaeToLoss,
                BenchmarkAlaeToLoss = item.BenchmarkAlaeToLoss,
                BenchmarkUnlimitedLossPlusAlaeRatio = item.BenchmarkUnlimitedLossPlusAlaeRatio,
                Frequency = item.Frequency * factor,
                LayerLossCostAmount = item.LayerLossCostAmount * factor,
                LayerLossCostPercent = item.LayerLossCostPercent,
                Severity = item.Severity,
                UnlimitedLossPlusAlaeRatio = item.UnlimitedLossPlusAlaeRatio,
            };
        }
    }
}