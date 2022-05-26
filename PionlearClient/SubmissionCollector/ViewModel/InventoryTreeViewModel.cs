using System.Collections.ObjectModel;
using SubmissionCollector.Models;

namespace SubmissionCollector.ViewModel
{
    internal class InventoryTreeViewModel : ViewModelBase
    {
        public ObservableCollection<IInventoryItem> PackageItems { get; set; }

        public InventoryTreeViewModel(IInventoryItem package)
        {
            var packageView = new ObservableCollection<IInventoryItem> { package };
            PackageItems = packageView;
        }
        
    }
}
