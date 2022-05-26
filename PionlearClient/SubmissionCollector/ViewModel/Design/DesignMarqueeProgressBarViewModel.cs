using System.Windows;
using SubmissionCollector.Models;
using SubmissionCollector.Properties;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignMarqueeProgressBarViewModel : BaseMarqueeProgressBarViewModel
    {
       public DesignMarqueeProgressBarViewModel()
        {
            Status = "This is status ...";
            Message = "This is a message";
            Image = Resources.Information.ToBitmapSource();
            StatusGridLength = new GridLength(1.0, GridUnitType.Star);
            MessageGridLength = new GridLength(1.0, GridUnitType.Star);
            ButtonsGridLength = ButtonHeight;
            BottomGridLength = BottomHeight;
            ShowZoom = true;
        }
    }
}
