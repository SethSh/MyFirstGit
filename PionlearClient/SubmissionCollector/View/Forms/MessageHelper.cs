using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Enums;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View.Forms
{
    public static class MessageHelper
    {
        private const bool DefaultFontSize = false;
        private const int SmallLineLength = 40;
        private const int SmallLineCount = 2;
        private const int MediumLineLength = 80;
        private const int MediumLineCount = 20;
        private const int LargeLineLength = 120;
        private const int LargeLineCount = 40;
        private const int SmallYesNoMessageThreshold = 50;

        public static void Show(string message)
        {
            Show(string.Empty, message, MessageType.Information);
        }

        public static void Show(string message, MessageType messageType)
        {
            Show(string.Empty, message, messageType);
        }

        public static void Show(string title, string message)
        {
            Show(title, message, MessageType.Information);
        }

        public static void Show(string title, string message, MessageType type)
        {
            var fullTitle = string.IsNullOrEmpty(title) ? BexConstants.ApplicationName : $"{BexConstants.ApplicationName}: {title}";

            var boxProperties = GetBoxProperties(message);
            var viewModel = new MessageBoxControlViewModel
            {
                Message = message,
                MessageType = type,
                ShowFontResizeAndExport = boxProperties.ShowFontResizeAndExport
            };
            
            var box = new MessageBoxControl(viewModel);
            var form = new MessageForm(box)
            {
                Height = (int) boxProperties.Height,
                Width = (int) boxProperties.Width,
                Text = fullTitle,
                StartPosition = FormStartPosition.CenterScreen
            };

            form.ShowDialog();
        }

        public static DialogResult ShowWithYesNo(string message)
        {
            var boxProperties = GetYesNoBoxProperties(message.Length);
            var yesNoBox = new MessageBoxYesNo(message, DefaultFontSize);
            var form = new MessageYesNoForm(yesNoBox)
            {
                Height = (int) boxProperties.Height,
                Width = (int) boxProperties.Width,
                Text = BexConstants.ApplicationName,
                StartPosition = FormStartPosition.CenterScreen,
            };
            form.ShowDialog();

            return yesNoBox.IsAnsweredYes ? DialogResult.Yes : DialogResult.No;
        }


        public static string ShowInputBox(string message, string defaultValue)
        {
            return Interaction.InputBox(message, BexConstants.ApplicationName, defaultValue);
        }

        public static BoxProperties GetBoxProperties(string message)
        {
            var boxProperties = new BoxProperties();

            var messageMetric = new MessageMetric();
            var messageMetrics = messageMetric.GetMessageMetrics(message).ToList();
            
            var lastLineWithText = messageMetrics.LastOrDefault(mms => mms.LineLength > 1);
            var lineCount = lastLineWithText?.LineNumber ?? 0;
            var maximumLineLength = messageMetrics.Max(mms => mms.LineLength);

            boxProperties.Height = GetHeight(lineCount);
            boxProperties.Width = GetWidth(maximumLineLength);
            
            boxProperties.ShowFontResizeAndExport = boxProperties.Height != FormSizeHeight.Small || boxProperties.Width != FormSizeWidth.Small;

            return boxProperties;
        }

        private static FormSizeHeight GetHeight(int lineCount)
        {
            if (lineCount <= SmallLineCount)
            {
                return FormSizeHeight.Small;
            }

            if (lineCount <= MediumLineCount)
            {
                return FormSizeHeight.Medium;
            }

            return lineCount <= LargeLineCount ? FormSizeHeight.Large : FormSizeHeight.ExtraLarge;
        }

        private static FormSizeWidth GetWidth(int lineLength)
        {
            if (lineLength <= SmallLineLength)
            {
                return FormSizeWidth.Small;
            }

            if (lineLength <= MediumLineLength)
            {
                return FormSizeWidth.Medium;
            }

            return lineLength <= LargeLineLength ? FormSizeWidth.Large : FormSizeWidth.ExtraLarge;
        }
        
        private static YesNoBoxProperties GetYesNoBoxProperties(int messageLength)
        {
            var boxProperties = new YesNoBoxProperties
            {
                Height = FormSizeHeight.Small,
                Width = messageLength <= SmallYesNoMessageThreshold ? FormSizeWidth.Small : FormSizeWidth.Medium
            };


            return boxProperties;
        }


    }

    public class YesNoBoxProperties
    {
        internal FormSizeHeight Height { get; set; }
        internal FormSizeWidth Width { get; set; }
    }

    public class BoxProperties: YesNoBoxProperties
    {
        internal bool ShowFontResizeAndExport { get; set; }
    }
}
