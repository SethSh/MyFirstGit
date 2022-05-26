using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Historicals.ExcelComponent;

namespace SubmissionCollector.Models.Historicals
{
    public class IndividualLossSet : BaseHistorical, IProvidesLedger
    {
        public IndividualLossSet(int segmentId, int componentId) : base(segmentId, componentId)
        {
            CommonExcelMatrix = new ExcelMatrix(SegmentId, ComponentId);
            Name = BexConstants.IndividualLossSetName;
            Ledger = new Ledger();
        }

        [JsonIgnore]
        public ExcelMatrix ExcelMatrix => (ExcelMatrix)CommonExcelMatrix;

        public Ledger Ledger { get; set; }
        
        public override void DecoupleFromServer()
        {
            base.DecoupleFromServer();
            Ledger.Clear();
        }

        protected override BaseSourceComponentModel MapToModel()
        {
            var isCombinedLossAndAlae = CommonExcelMatrix.GetSegment().IndividualLossSetDescriptor.IsLossAndAlaeCombined;

            return new IndividualLossSetModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                Id = ComponentId,
                Guid = Guid,
                Name = ExcelMatrix.FullName,
                IsCombinedLossAndAlae = isCombinedLossAndAlae,
                Threshold = ExcelMatrix.Threshold,
                SublineIds = ExcelMatrix.Select(x => new long?(x.Code)).ToList(),
                Items = ExcelMatrix.Items,
                InterDisplayOrder = ExcelMatrix.InterDisplayOrder,
                IntraDisplayOrder = ExcelMatrix.IntraDisplayOrder
            };
        }
    }
}
