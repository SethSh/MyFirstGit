using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using PionlearClient.Model;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Enums;
using SubmissionCollector.Models.Package.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Package
{
    public interface IPackage : IModel 
    {
        List<IExcelMatrix> ExcelMatrices { get; set; }
        StringBuilder ExcelValidation { get; set; }
        PackageModel CreatePackageModel();
        IList<BexCommunicationEntry> BexCommunications { get; set; }
        AttachOptions AttachedToRatingAnalysisStatus { get; }
        string AttachedToRatingAnalysisMessage { get; }
        void AddItemToBexCommunication(string activity);
        ObservableCollection<ISegment> Segments { get; set; }
        bool HasSourceId { get; }
        string ResponsibleUnderwriterId { get; set; }
        string ResponsibleUnderwriter { get; set; }
        string CedentId { get; set; }
        string CedentName { get; set; }
        UnderwritingYearExcelMatrix UnderwritingYearExcelMatrix { get; }
        bool IsSelected { get; set; }
        void SetAllItemsToNotDirty();
    }
}
