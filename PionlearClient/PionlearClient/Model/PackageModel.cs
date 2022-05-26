using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.Extensions;

namespace PionlearClient.Model
{
    public class PackageModel : BaseSourceModel
    {
        private const int PackageNameMaxLength = 200;
        private const string NewPackageName = "A New Package";
        private const string NewCedentName = "A  New Company"; // this double space is not a typo but a MR convention
        private const string NewCedentId = "099999";
        private const string UsCurrencyCode = "USD";

        public PackageModel()
        {
            AsOfDate = null;
            UnderwritingYear = string.Empty;
            Name = NewPackageName;
            CedentName = NewCedentName;
            CedentId = NewCedentId;
            Currency = UsCurrencyCode;
            Guid = Guid.NewGuid();
            IsDirty = true;
            SegmentModels = new List<SegmentModel>();
        }

        public string CedentId { get; set; }
        public DateTime? AsOfDate { get; set; }
        public string UnderwritingYear { get; set; }
        public string AnalystId { get; set; }
        public string Currency { get; set; }
        public List<SegmentModel> SegmentModels { get; set; }
        public string CedentName { get; set; }

        public SubmissionPackageModel Map()
        {
            var asOfDate = DateTime.Now.Date;
            if (AsOfDate.HasValue) asOfDate = AsOfDate.Value.Date;

            var submissionPackageModel = new SubmissionPackageModel
            {
                Id = SourceId,
                AsOfDate = asOfDate,
                CedentId = CedentId,
                Currency = Currency,
                Name = Name,
                AnalystId = AnalystId,
                UnderwritingYear = UnderwritingYear
            };

            return submissionPackageModel;
        }
        
        public StringBuilder PerformQualityControl()
        {
            var messages = new StringBuilder();
            
            if (CedentId == BexConstants.DefaultCedentId)
            {
                messages.AppendLine(Name);
                messages.AppendLine($"Change {BexConstants.CedentName.ToLower()} from default {BexConstants.CedentName.ToLower()} <{BexConstants.DefaultCedentId}>.");
                messages.AppendLine(string.Empty);
            }

            foreach (var segmentModel in SegmentModels)
            {
                var segmentQc = segmentModel.PerformQualityControl();
                if (segmentQc.Length > 0)
                {
                    messages.AppendLine(segmentModel.Name);
                    messages.Append(segmentQc);
                    messages.AppendLine(string.Empty);
                }

                //why is this here
                segmentModel.IndividualLossSetModels.ForEach(model =>
                    model.SubjectPolicyAlaeTreatment = segmentModel.SubjectPolicyAlaeTreatment);

                foreach (var model in segmentModel.Models
                    .OrderBy(model => model.InterDisplayOrder)
                    .ThenBy(model => model.IntraDisplayOrder))
                {
                    model.BuildQualityControlMessage(messages, segmentModel.Name);
                }
                
            }
            return messages;
        }
        
        public StringBuilder Validate()
        {
            var messages = new StringBuilder();

            var packageValidation = ValidatePackageOnly();
            if (packageValidation.Length > 0)
            {
                messages.AppendLine($"{BexConstants.PackageName}: {Name}");
                messages.AppendLine(packageValidation.ToString());
                messages.AppendLine(string.Empty);
            }

            var duplicateSegmentNames = SegmentModels.Select(s => s.Name).Duplicates().ToList();
            if (duplicateSegmentNames.Any())
            {
                var duplicatesString = string.Join(", ", duplicateSegmentNames.ToArray());
                var message = $"{BexConstants.SegmentName} can't contain duplicate names.  The following names appear more than once: {duplicatesString})";
                messages.AppendLine(message);
            }

            foreach (var segmentModel in SegmentModels)
            {
                var riskClassValidation = segmentModel.Validate();
                if (riskClassValidation.Length > 0)
                {
                    messages.AppendLine(segmentModel.Name);
                    messages.AppendLine(riskClassValidation.ToString());
                    messages.AppendLine(string.Empty);
                }

                foreach (var model in segmentModel.Models
                    .OrderBy(model => model.InterDisplayOrder)
                    .ThenBy(model => model.IntraDisplayOrder))
                {
                    model.BuildValidationMessage(segmentModel.ValidSublines, messages, segmentModel.Name);
                }
            }
            return messages;
        }

        
        private StringBuilder ValidatePackageOnly()
        {
            var validation = new StringBuilder();
            if (string.IsNullOrEmpty(Name))
            {
                validation.AppendLine($"{BexConstants.PackageName.ToStartOfSentence()} name can't be blank");
            }
            else
            {
                if (Name.Length > PackageNameMaxLength)
                    validation.AppendLine($"{BexConstants.PackageName.ToStartOfSentence()} name length can't be greater than {PackageNameMaxLength} characters");

#if !DEBUG
                    if (Name == NewPackageName) validation.AppendLine($"{BexConstants.PackageName.ToStartOfSentence()} name can't be left as the default value");
#endif
            }

            if (string.IsNullOrEmpty(CedentId)) validation.AppendLine($"{BexConstants.CedentName} can't be blank");
            if (AsOfDate == null) validation.AppendLine("As of date can't be blank");
            if (string.IsNullOrEmpty(UnderwritingYear)) validation.AppendLine($"{BexConstants.UnderwritingYearName.ToStartOfSentence()} can't be blank");
            if (string.IsNullOrEmpty(AnalystId)) validation.AppendLine($"{BexConstants.AnalystName} can't be blank");

            return validation;
        }
    }
}
