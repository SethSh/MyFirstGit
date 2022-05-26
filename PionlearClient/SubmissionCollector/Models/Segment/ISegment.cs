using System.Collections.Generic;
using System.Text;
using PionlearClient.Model;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Models.Profiles;
using SubmissionCollector.Models.Profiles.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Models.Segment
{
    public interface ISegment : ICollection<ISubline>, IModelProxy, IInventoryItem
    {
        new string Name { get; set; }
        int Id { get; set; }
        string ProspectiveExposureBasis { get; set; }
        ProspectiveExposureAmountExcelMatrix ProspectiveExposureAmountExcelMatrix { get; }
        SegmentModel SegmentModel { get; set; }
        bool IsUmbrella { get; set; }
        IPackage ParentPackage { get; } 

        string HistoricalPeriodType { get; set; }
        string HistoricalExposureBasis { get; set; }
        string ProspectiveExposureAmountExcelCell { get; }

        
        List<IExcelComponent> ExcelComponents { get; set; }
        IList<ISegmentExcelMatrix> ExcelMatrices { get; }
        IList<ISegmentExcelMatrix> OrphanExcelMatrices { get; set; }
        IAggregateLossSetDescriptor AggregateLossSetDescriptor { get; set; }
        IIndividualLossSetDescriptor IndividualLossSetDescriptor { get; set; }
        
        SublineExcelMatrix SublineExcelMatrix { get; }
        UmbrellaExcelMatrix UmbrellaExcelMatrix { get; }

        IEnumerable<IProfile> Profiles { get; }
        IEnumerable<IHistorical> Historicals { get; }

        IEnumerable<PolicyProfile> PolicyProfiles { get; }

        IEnumerable<StateProfile> StateProfiles { get; }
        IEnumerable<HazardProfile> HazardProfiles { get; }

        IEnumerable<ConstructionTypeProfile> ConstructionTypeProfiles { get; }
        IEnumerable<OccupancyTypeProfile> OccupancyTypeProfiles { get; }
        IEnumerable<ProtectionClassProfile> ProtectionClassProfiles { get; }
        IEnumerable<TotalInsuredValueProfile> TotalInsuredValueProfiles { get; }

        WorkersCompStateHazardGroupProfile WorkersCompStateHazardGroupProfile { get; }
        WorkersCompClassCodeProfile WorkersCompClassCodeProfile { get; }
        WorkersCompStateAttachmentProfile WorkersCompStateAttachmentProfile { get; }
        MinnesotaRetention WorkersCompMinnesotaRetention { get; }
    

        PeriodSet PeriodSet { get; }
        IEnumerable<ExposureSet> ExposureSets { get; }
        IEnumerable<AggregateLossSet> AggregateLossSets { get; }
        IEnumerable<IndividualLossSet> IndividualLossSets { get; }
        IEnumerable<RateChangeSet> RateChangeSets { get; }

        WorksheetManager WorksheetManager { get; set; }
        string HeaderRangeName { get; }

        bool ContainsAnyCommercialSublines { get; }
        bool ContainsAnyPersonalSublines { get; }
        
        bool IsLiability { get; }
        bool IsProperty { get; }
        bool IsWorkersComp { get; }

        bool IsStructureModifiable { get; }
        bool AreUmbrellaTypesModifiable { get; }
        string PolicyLimitsApplyTo { get; set; }
        SublineProfile SublineProfile { get; }
        bool IsCurrentlyRebuilding { get; set; }
        bool IsWorkerCompClassCodeActive { get; set; }
        string GetBlockModificationsMessage(string type);

        ISegment Duplicate();

        void UpdateSublines(IList<ISubline> changedSublines);
        
        void AddConstructionTypeProfile(ISubline item);
        void AddProtectionClassProfile(ISubline item);
        void AddTotalInsuredValueProfile(ISubline item);
        void AddOccupancyTypeProfile(ISubline item);

        void AddWorkersCompClassCodeProfile();
        void AddWorkersCompStateAttachmentProfile();
        void AddWorkersCompStateHazardGroupProfile();
        void AddWorkersCompMinnesotaRetention();
        

        StringBuilder CreateSegmentModel();
        StringBuilder CreateSegmentModelForQualityControl();
        
        void RegisterBaseChangeEvents();
        bool IsNameAcceptableLength();
        void SetAllLedgerItemsToNotDirty();
    }
}