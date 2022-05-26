using PionlearClient.KeyDataFolder;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;

namespace SubmissionCollector.ViewModel.ItemSources
{
    public class CurrencySource : IItemsSource
    {
        public ItemCollection GetValues()
        {
            var currencyChoices = new ItemCollection();
            foreach (var item in CurrenciesFromKeyData.CurrencyReferenceData)
            {
                currencyChoices.Add(item.IsoCode, item.Name);
            }
            return currencyChoices;
        }
    }
}
