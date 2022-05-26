using System.Windows;
using System.Windows.Media.Imaging;

namespace SubmissionCollector.ViewModel
{
    public interface IMarqueeProgressBarViewModel
    {
        string Status { get; set; }
        string Message { get; set; }
        BitmapSource Image { get; set; }
    }

    public class MarqueeProgressBarViewModel : BaseMarqueeProgressBarViewModel
    {
        
    }

    public abstract class BaseMarqueeProgressBarViewModel : ViewModelBase, IMarqueeProgressBarViewModel
    {
        private string _status;
        private string _message;
        private BitmapSource _image;
        private GridLength _buttonsRowPixels;
        private bool _showZoom;
        private GridLength _messageRowPixels;
        private GridLength _statusRowPixels;
        protected GridLength ButtonHeight = new GridLength(40);
        protected GridLength BottomHeight = new GridLength(40);
        private GridLength _bottomRowPixels;

        protected BaseMarqueeProgressBarViewModel()
        {
            SetAppearanceToStatus();
        }

        public void SetAppearanceToStatus()
        {
            Status = "Starting ...";
            StatusGridLength = new GridLength(1, GridUnitType.Star);
            MessageGridLength = new GridLength(0);
            ButtonsGridLength = new GridLength(0);
            BottomGridLength = BottomHeight;
        }

        public void SetFailureAppearance()
        {
            SetCommonAppearance();
            ShowZoom = true;
        }

        public void SetSuccessAppearance()
        {
            SetCommonAppearance();
            ShowZoom = false;
        }

        private void SetCommonAppearance()
        {
            StatusGridLength = new GridLength(0, GridUnitType.Star);
            MessageGridLength = new GridLength(1, GridUnitType.Star);
            ButtonsGridLength = ButtonHeight;
            BottomGridLength = new GridLength(0);
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

        public BitmapSource Image
        {
            get => _image;
            set
            {
                _image = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength MessageGridLength
        {
            get => _messageRowPixels;
            set
            {
                _messageRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength ButtonsGridLength
        {
            get => _buttonsRowPixels;
            set
            {
                _buttonsRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength BottomGridLength
        {
            get => _bottomRowPixels;
            set
            {
                _bottomRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public bool ShowZoom
        {
            get => _showZoom;
            set
            {
                _showZoom = value;
                NotifyPropertyChanged();
            }
        }

        public GridLength StatusGridLength
        {
            get => _statusRowPixels;
            set
            {
                _statusRowPixels = value;
                NotifyPropertyChanged();
            }
        }

        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                NotifyPropertyChanged();
            }
        }

        
    }
}
