using System;

namespace SubmissionCollector.ViewModel.Design
{
    public class DesignHistoryViewModel : BaseHistoryViewModel
    {
        public DesignHistoryViewModel()
        {
            Items.Add(new HistoryItemViewModel {UserName = "Seth Shenghit", Timestamp = DateTime.Now, Activity = "None"});
            Items.Add(new HistoryItemViewModel { UserName = "Jim Sandor", Timestamp = new DateTime(2020, 12, 22), Activity = "Eating cake" });
        }
    }
}