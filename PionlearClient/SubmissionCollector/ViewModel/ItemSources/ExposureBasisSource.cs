using PionlearClient.BexReferenceData;
using PionlearClient.Extensions;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace SubmissionCollector.ViewModel.ItemSources
{
    public class ExposureBasisSource : IItemsSource
    {
        public ItemCollection GetValues()
        {
            var exposureUnits = new ItemCollection();
            ExposureBasisFromBex.ReferenceData.ForEach(x => exposureUnits.Add(x.Id, x.ExposureBaseName));
            return exposureUnits;
        }
    }
}
