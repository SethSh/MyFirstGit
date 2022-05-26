﻿using Newtonsoft.Json;
using PionlearClient.Model;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.Models.Profiles
{
    public class UmbrellaProfile : BaseProfile
    {
        public UmbrellaProfile(int segmentId) : base(segmentId)
        {
            CommonExcelMatrix = new UmbrellaExcelMatrix(segmentId);
        }

        [JsonIgnore]
        public UmbrellaExcelMatrix UmbrellaExcelMatrix => (UmbrellaExcelMatrix)CommonExcelMatrix;

        public override bool IsDirty
        {
            get => base.IsDirty;
            set
            {
                var excelWorkspace = Globals.ThisWorkbook.ThisExcelWorkspace;
                if (excelWorkspace == null) return;

                var segment = excelWorkspace.Package.GetSegment(SegmentId);
                if (segment != null && value) segment.IsDirty = true;

                base.IsDirty = value;
            }
        }

        protected override BaseSourceComponentModel MapToModel()
        {
            throw new System.NotImplementedException();
        }
    }
}
