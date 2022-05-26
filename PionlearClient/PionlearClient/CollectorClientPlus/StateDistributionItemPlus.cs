using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class StateDistributionItemPlus : StateDistributionItem, IProvidesLocation
    {
        public string Location { get; set; }
    }
}
