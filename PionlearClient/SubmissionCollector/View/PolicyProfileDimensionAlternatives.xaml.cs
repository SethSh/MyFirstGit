using System.Windows;
using System.Windows.Controls;
using PionlearClient;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for MessageBoxYesNo.xaml
    /// </summary>
    public partial class PolicyProfileDimensionAlternatives
    {
        private const int StartFontSize = 14;

        public PolicyProfileDimensionAlternatives(string message)
        {
            InitializeComponent();

            MyLabel.Content = $"{BexConstants.PolicyProfileName} Dimension Alternatives";
            MyTextBlock.Text = message;
            CloseButton.Focus();
            FontSizeSlider.Value = StartFontSize;
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
