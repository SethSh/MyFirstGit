using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class MinnesotaRetention : BaseProfile
    {
        public MinnesotaRetention(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new MinnesotaRetentionExcelMatrix(segmentId);
            Name = BexConstants.MinnesotaRetentionName;
        }

        [JsonIgnore]
        public MinnesotaRetentionExcelMatrix ExcelMatrix => (MinnesotaRetentionExcelMatrix)CommonExcelMatrix;

        protected override BaseSourceComponentModel MapToModel()
        {
            return new MinnesotaRetentionModel
            {
                RetentionId = ExcelMatrix.RetentionId
            };
        }
    }
}
