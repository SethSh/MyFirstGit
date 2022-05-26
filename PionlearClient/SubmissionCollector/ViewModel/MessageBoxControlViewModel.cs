using SubmissionCollector.Enums;

namespace SubmissionCollector.ViewModel
{
    public class MessageBoxControlViewModel : BaseMessageBoxControlViewModel
    {
        public MessageBoxControlViewModel()
        {
            ShowFontResizeAndExport = true;
        }
    }

    public abstract class BaseMessageBoxControlViewModel : ViewModelBase, IMessageBoxControlViewModel
    {
        private string _message;
        private MessageType _messageType;
        private bool _showFontResizeAndExport;
        
        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                NotifyPropertyChanged();
            }
        }

        public MessageType MessageType
        {
            get => _messageType;
            set
            {
                _messageType = value;
                NotifyPropertyChanged();
            }
        }

        public bool ShowFontResizeAndExport
        {
            get => _showFontResizeAndExport;
            set
            {
                _showFontResizeAndExport = value;
                NotifyPropertyChanged();
            }
        }
    }

    public interface IMessageBoxControlViewModel
    {
        string Message { get; set; }
        MessageType MessageType { get; set; }
        bool ShowFontResizeAndExport { get; set; }
    }
}