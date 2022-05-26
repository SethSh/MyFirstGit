using System.Windows;
using System.Windows.Controls;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for ScrollableYesNo.xaml
    /// </summary>
    public partial class MessageBoxYesNo
    {
        public bool IsAnsweredYes;
        private const int StartFontSize = 14;

        public MessageBoxYesNo(string message, bool showFontResize)
        {
            InitializeComponent();
            MyTextBlock.Text = message;
            IsAnsweredYes = false;
            YesButton.Focus();

            var vis = showFontResize ? Visibility.Visible : Visibility.Hidden;
            FontSizeSlider.Visibility = vis;
            FontSizeSlider.Value = StartFontSize;
        }

        private void YesButton_OnClick(object sender, RoutedEventArgs e)
        {
            IsAnsweredYes = true;
        }
        
        private void NoButton_OnClick(object sender, RoutedEventArgs e)
        {
            IsAnsweredYes = false;
        }


        private void FontSizeSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var slider = sender as Slider;
            if (slider == null) return;

            var value = slider.Value;
            MyTextBlock.FontSize = value;
        }
    }
}
