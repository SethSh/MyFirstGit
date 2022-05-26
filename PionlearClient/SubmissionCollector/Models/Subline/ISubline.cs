using System.Windows.Media.Imaging;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Subline
{
    public interface ISubline : IInventoryItem
    {
        int SegmentId { get; set; }
        string ShortName { get;  }
        string ShortNameWithLob { get; }
        string NameWithLob { get; }
        string LobShortName { get; }
        bool HasPolicyProfile { get; }
        bool HasStateProfile { get;  }
        bool HasHazardProfile { get; }
        bool IsLineExclusive { get; }
        int Code { get;  }
        BitmapSource ImageSource { get; set; }
        bool IsPersonal { get; }
        ISegment FindParentSegment();
        bool IsWorkersComp { get; }
        LineOfBusinessType LineOfBusinessType { get; set; }
    }

    public enum LineOfBusinessType
    {
        Property,
        Liability,
        WorkersCompensation,
        PhysicalDamage
    }
}
