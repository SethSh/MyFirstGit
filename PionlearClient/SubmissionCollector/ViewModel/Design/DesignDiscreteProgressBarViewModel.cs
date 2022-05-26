using System.Windows;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignDiscreteProgressBarViewModel : IDiscreteProgressBarViewModel
    {
        public double DoneAmount { get; set; }
        public string Message { get; set; }
        public GridLength ButtonsRowPixels { get; set; }

        public DesignDiscreteProgressBarViewModel()
        {
            DoneAmount = 50;
            Message = "Message goes here";
            ButtonsRowPixels = new GridLength(40);
        }
    }
}
