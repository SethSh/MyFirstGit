using SubmissionCollector.Enums;

namespace SubmissionCollector.ViewModel.Design
{
    internal class DesignMessageBoxControlViewModel : BaseMessageBoxControlViewModel
    {
        public DesignMessageBoxControlViewModel()
        {
            Message = "This is a test message. There is a lot of things to say to get to the end of the first line.";
            MessageType = MessageType.Documentation;
            ShowFontResizeAndExport = true;
        }
    }
}
