using System.Collections.Generic;
using System.Linq;
using System.Windows;
using PionlearClient.KeyDataFolder;

namespace SubmissionCollector.ViewModel.Design
{
    internal sealed class DesignUnderwriterSelectorViewModel : BaseUnderwriterSelectorViewModel
    {
        public DesignUnderwriterSelectorViewModel()
        {
            Criteria = "Jim Sandor";
            UnderwriterCount = 25;
            ShowMyUnderwriters = true;
            CriteriaRowLength = new GridLength(40);
            Underwriters = new List<Underwriter>
            {
                new Underwriter {Name = "The Cat In the Hat", Code = "N1234567"} ,
                new Underwriter {Name = "Mickey Mouse", Code = "N1000000"} ,
                new Underwriter {Name = "Felix the Cat", Code = "N1111111"} ,
            };
            FilteredUnderwriters = Underwriters.Where(underwriter => underwriter.Name.Contains("Cat")).ToList();
        }
    }
}
