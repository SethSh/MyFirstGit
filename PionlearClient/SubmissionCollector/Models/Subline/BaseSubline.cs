using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Media.Imaging;
using Newtonsoft.Json;
using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Subline
{
    [Serializable]
    public abstract class BaseSubline : ISubline
    {
        protected BaseSubline()
        {
            // ReSharper disable once VirtualMemberCallInConstructor
            LineOfBusinessType = LineOfBusinessType.Liability;
        }

        [Browsable(false)]
        public bool IsLineExclusive => !SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).LineOfBusiness.CanBeCombinedWithSiblingLobs;

        [Browsable(false)]
        public bool HasPolicyProfile => SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).LineOfBusiness.IncludesPolicyLimitProfile;
        
        [Browsable(false)]
        public bool HasStateProfile => !SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).HasDefaultState;

        [Browsable(false)]
        public virtual bool HasHazardProfile => !SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).HasDefaultHazard;

        [Browsable(false)]
        public int Code { get; set; }

        [Browsable(false)]
        public int SegmentId { get; set; }

        [Browsable(false)]
        public int DisplayOrder { get; set; }

        [Browsable(false)]
        public bool IsSelected { get; set; }

        [Browsable(false)]
        public bool IsExpanded { get; set; }

    
        [JsonIgnore]
        [Browsable(false)]
        public BitmapSource ImageSource { get; set; }

        [Browsable(false)]
        public string Name => SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).SublineName;

        [JsonIgnore]
        [Browsable(false)]
        public string NameWithLob
        {
            get
            {
                var s = SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code);
                return $"{s.LineOfBusiness.Name.ConnectWithDash(s.SublineName)}";
            }
        }
        
        [JsonIgnore]
        [Browsable(false)]
        public string ShortName => SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).SublineShortName;
        
        [JsonIgnore]
        [Browsable(false)]
        public string LobShortName
        {
            get
            {
                return SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).LineOfBusiness.ShortName;
            }
        }

        [JsonIgnore]
        [Browsable(false)]
        public string ShortNameWithLob => LobShortName.ConnectWithDash(ShortName);

        [JsonIgnore]
        [Browsable(false)]
        public bool IsPersonal => SublineCodesFromBex.ReferenceData.Single(x => x.SublineId == Code).IsPersonal;

        [JsonIgnore]
        [Browsable(false)]
        public bool IsWorkersComp => this.LineOfBusinessType == LineOfBusinessType.WorkersCompensation;


        public ISegment FindParentSegment()
        {
            return Globals.ThisWorkbook.ThisExcelWorkspace.Package.Segments.Single(segment => segment.Id == SegmentId);
        }

        [Browsable(false)]
        public LineOfBusinessType LineOfBusinessType { get; set; }
        
    }
}
