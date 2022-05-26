using System.Windows;

namespace SubmissionCollector.ViewModel
{
    internal interface IDiscreteProgressBarViewModel
    {
        string Message { get; set; }
    }

    public class DiscreteProgressBarViewModel : ViewModelBase, IDiscreteProgressBarViewModel
    {
        private string _message;
        private double _donePercent;
        private GridLength _buttonsRowPixels;
        public static GridLength ButtonRowsPixelsCollapsed = new GridLength(0);
        public static GridLength ButtonRowsPixelsExpanded = new GridLength(40);

        public double DoneAmount
        {
            get => _donePercent;
            set
            {
                _donePercent = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength ButtonsRowPixels
        {
            get => _buttonsRowPixels;
            set
            {
                _buttonsRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                NotifyPropertyChanged();
            }
        }
    }
}
