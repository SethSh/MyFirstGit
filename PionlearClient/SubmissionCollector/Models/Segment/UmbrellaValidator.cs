using System;
using System.Collections.Generic;
using System.Linq;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models.Segment
{
    internal class UmbrellaValidator
    {
        public bool Validate(ISegment segment)
        {
            if (!ValidateIsUmbrella(segment)) return false;

            segment.UmbrellaExcelMatrix.Validate();

            var umbrellaResult = ValidateAllocationCount(segment);
            if (!umbrellaResult.IsValid) return false;

            var orphanUmbrellaTypeCodes = umbrellaResult.UmbrellaTypeCodeInUse.Except(umbrellaResult.UmbrellaAllocationCodes).ToList();
            if (!orphanUmbrellaTypeCodes.Any()) return true;

            var orphanUmbrellaTypeNames = UmbrellaTypesFromBex.GetNames(orphanUmbrellaTypeCodes).ToList();
            var orphanMessageStart = $"These {BexConstants.UmbrellaTypeName.ToLower()}s exist in a {BexConstants.PolicyProfileName.ToLower()} " +
                                     $"but don't have values in the {BexConstants.UmbrellaAllocationName.ToLower()}.";

            var orphanMessageEnd =
                $"The {BexConstants.UmbrellaTypeName.ToLower()} wizard requires all of the {BexConstants.UmbrellaTypeName.ToLower()}s specified " +
                $"in the {BexConstants.PolicyProfileName}s to have values " +
                $"in the {BexConstants.UmbrellaAllocationName.ToLower()}.";

            var message = $"{segment.Name}: {orphanMessageStart} \n\n\t{string.Join("\n\t", orphanUmbrellaTypeNames)} \n\n{orphanMessageEnd}";
            MessageHelper.Show(message, MessageType.Stop);

            return false;
        }

        private static bool ValidateIsUmbrella(ISegment segment)
        {
            if (segment.IsUmbrella) return true;

            MessageHelper.Show($"{segment.Name}: umbrella has not been selected", MessageType.Stop);
            return false;
        }

        private static UmbrellaResult ValidateAllocationCount(ISegment segment)
        {
            var commercialUmbrellaTypeCodeInUse = segment.PolicyProfiles
                .Where(x => x.UmbrellaType.HasValue && x.UmbrellaType.Value != UmbrellaTypesFromBex.GetPersonalCode())
                .Select(x => x.UmbrellaType.Value).ToList();
            
            var umbrellaAllocationCodes = segment.UmbrellaExcelMatrix.Allocations.Select(x => Convert.ToInt32(x.Id));
            if (segment.UmbrellaExcelMatrix.Allocations.Count > 1) return new UmbrellaResult
            {
                UmbrellaAllocationCodes =  umbrellaAllocationCodes,
                UmbrellaTypeCodeInUse = commercialUmbrellaTypeCodeInUse,
                IsValid = true
            };

            var notReadyMessage = $"The {BexConstants.UmbrellaTypeName.ToLower()} wizard requires at least two populated rows in " +
                                  $"the {BexConstants.UmbrellaAllocationName.ToLower()}.";

            MessageHelper.Show($"{segment.Name}: {notReadyMessage}", MessageType.Stop);
            return new UmbrellaResult { IsValid = false };
        }
        
        internal struct UmbrellaResult
        {
            internal List<int> UmbrellaTypeCodeInUse;
            internal IEnumerable<int> UmbrellaAllocationCodes;
            internal bool IsValid;
        }
    }
}
