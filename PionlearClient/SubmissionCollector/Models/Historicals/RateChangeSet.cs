using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using RateChangeSetModel = PionlearClient.Model.RateChangeSetModel;

namespace SubmissionCollector.Models.Historicals
{
    public class RateChangeSet : BaseHistorical, IProvidesLedger
    {
        public RateChangeSet(int segmentId, int componentId) : base(segmentId, componentId)
        {
            CommonExcelMatrix = new RateChangeExcelMatrix(SegmentId, ComponentId);
            Name = BexConstants.RateChangeName;
            Ledger = new Ledger();
        }
        
        [JsonIgnore]
        public RateChangeExcelMatrix ExcelMatrix => (RateChangeExcelMatrix)CommonExcelMatrix;

        public Ledger Ledger { get; set; }

        public override void DecoupleFromServer()
        {
            base.DecoupleFromServer();
            Ledger.Clear();
        }

        protected override BaseSourceComponentModel MapToModel()
        {
            return new RateChangeSetModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                Id = ComponentId,
                Guid = Guid,
                Name = ExcelMatrix.FullName,
                SublineIds = ExcelMatrix.Select(x => new long?(x.Code)).ToList(),
                Items = ExcelMatrix.Items,
                InterDisplayOrder = ExcelMatrix.InterDisplayOrder,
                IntraDisplayOrder = ExcelMatrix.IntraDisplayOrder
            };
        }
    }
}
