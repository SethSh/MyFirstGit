using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace SubmissionCollector.ViewModel.ItemSources
{
    internal class HistoricalPeriodTypeSource : IItemsSource
    {
        public ItemCollection GetValues()
        {
            var itemCollection = new ItemCollection();
            HistoricalPeriodTypesFromBex.ReferenceData.ForEach(x => itemCollection.Add(x.Id, x.Name));
            return itemCollection;
        }
    }
}
