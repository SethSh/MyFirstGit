using System.Collections.Generic;
using System.Windows;
using PionlearClient.KeyDataFolder;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignCedentSelectorViewModel : BaseCedentSelectorViewModel
    {
        public DesignCedentSelectorViewModel()
        {
            Criteria = "AIG";
            StatusMessage = "This is running ...";
            ShowMyCedents = true;
            CriteriaRowLength = new GridLength(40);
            StatusRowLength = new GridLength(40);
            CedentCount = 25;

            Cedents = new List<BusinessPartner>
            {
                new BusinessPartner
                {
                    City = "Brooklyn",
                    Id = "0000123456",
                    Name = "My Favorite cedent"
                }
            };
        }
    }
}
