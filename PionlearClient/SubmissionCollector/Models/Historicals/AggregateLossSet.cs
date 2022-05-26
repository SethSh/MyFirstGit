using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Historicals.ExcelComponent;

namespace SubmissionCollector.Models.Historicals
{
    public class AggregateLossSet : BaseHistorical, IProvidesLedger
    {
        public AggregateLossSet(int segmentId, int componentId) : base(segmentId, componentId)
        {
            CommonExcelMatrix = new AggregateLossSetExcelMatrix(SegmentId, ComponentId);
            Name = BexConstants.AggregateLossSetName;
            Ledger = new Ledger();
        }

        [JsonIgnore]
        public AggregateLossSetExcelMatrix ExcelMatrix => (AggregateLossSetExcelMatrix)CommonExcelMatrix;

        public Ledger Ledger { get; set; }

        public override void DecoupleFromServer()
        {
            base.DecoupleFromServer();
            Ledger.Clear();
        }

        protected override BaseSourceComponentModel MapToModel()
        {
            var aggregateLossSetDescriptor = CommonExcelMatrix.GetSegment().AggregateLossSetDescriptor;

            return new AggregateLossSetModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                Id = ComponentId,
                Guid = Guid,
                IsCombinedLossAndAlae = aggregateLossSetDescriptor.IsLossAndAlaeCombined,
                IsPaidAvailable = aggregateLossSetDescriptor.IsPaidAvailable,
                SublineIds = ExcelMatrix.Select(x => new long?(x.Code)).ToList(),
                Items = ExcelMatrix.Items,
                Name = ExcelMatrix.FullName,
                InterDisplayOrder = ExcelMatrix.InterDisplayOrder,
                IntraDisplayOrder = ExcelMatrix.IntraDisplayOrder
            };
        }
    }
}
