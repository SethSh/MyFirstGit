namespace SubmissionCollector.ViewModel
{
    public class MessageBoxForFactorsViewModel : BaseMessageBoxForFactorsViewModel
    {
        public MessageBoxForFactorsViewModel(string message, string renameMessage, string replaceMessage)
        {
            Message = message;
            RenameMessage = renameMessage;
            ReplaceMessage = replaceMessage;
            UpdateFactorOption = UpdateFactorOption.Cancel;
        }
    }

    public abstract class BaseMessageBoxForFactorsViewModel : ViewModelBase, IMessageBoxForFactorsViewModel
    {
        private string _message;
        private UpdateFactorOption _updateFactorOption;
        private string _renameMessage;
        private string _replaceMessage;

        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                NotifyPropertyChanged();
            }
        }

        public string RenameMessage
        {
            get => _renameMessage;
            set
            {
                _renameMessage = value;
                NotifyPropertyChanged();
            }
        }

        public string ReplaceMessage
        {
            get => _replaceMessage;
            set
            {
                _replaceMessage = value;
                NotifyPropertyChanged();
            }
        }

        public UpdateFactorOption UpdateFactorOption
        {
            get => _updateFactorOption;
            set
            {
                _updateFactorOption = value;
                NotifyPropertyChanged();
            }
        }
    }

    public interface IMessageBoxForFactorsViewModel
    {
        string Message { get; set; }
        UpdateFactorOption UpdateFactorOption { get; set; }
    }

    public enum UpdateFactorOption
    {
        Rename,
        Delete,
        Cancel,
        None
    }
}
