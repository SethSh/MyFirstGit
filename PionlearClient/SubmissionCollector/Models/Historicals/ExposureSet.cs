using System;
using System.Linq;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.Models.Historicals.ExcelComponent;

namespace SubmissionCollector.Models.Historicals
{
    public class ExposureSet : BaseHistorical, IProvidesLedger
    {
        public ExposureSet(int segmentId, int componentId) : base (segmentId, componentId)
        {
            CommonExcelMatrix = new ExposureSetExcelMatrix(SegmentId, ComponentId);
            Name = BexConstants.ExposureSetName;
            Ledger = new Ledger();
        }

        [JsonIgnore]
        public ExposureSetExcelMatrix ExcelMatrix => (ExposureSetExcelMatrix)CommonExcelMatrix;

        public Ledger Ledger { get; set; }

        public override void DecoupleFromServer()
        {
            base.DecoupleFromServer();
            Ledger.Clear();
        }

        protected override BaseSourceComponentModel MapToModel()
        {
            var historicalExposureBasis = CommonExcelMatrix.GetSegment().HistoricalExposureBasis;

            return new ExposureSetModel
            {
                IsDirty = IsDirty,
                SourceId = SourceId,
                Id = ComponentId,
                Guid = Guid,
                SublineIds = ExcelMatrix.Select(x => new long?(x.Code)).ToList(),
                Items = ExcelMatrix.Items,
                Name = ExcelMatrix.FullName,
                ExposureBaseId = Convert.ToInt16(historicalExposureBasis),
                InterDisplayOrder = ExcelMatrix.InterDisplayOrder,
                IntraDisplayOrder = ExcelMatrix.IntraDisplayOrder
            };
        }
    }
}
