using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace SubmissionCollector.ViewModel.ItemSources
{
    public class PolicyLimitsApplyToSource : IItemsSource
    {
        public ItemCollection GetValues()
        {
            var policyLimitsApplyToChoices = new ItemCollection();
            SubjectPolicyAlaeTreatmentsFromBex.ReferenceData.ForEach(x => policyLimitsApplyToChoices.Add(x.Id, x.Description));
            return policyLimitsApplyToChoices;
        }
    }
}