using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using PionlearClient.BexReferenceData;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for IsDirtyFlagDisplay.xaml
    /// </summary>
    public partial class IsDirtyFlagDisplay
    {
        private const int StartFontSize = 14;

        public IsDirtyFlagDisplay(string message)
        {
            InitializeComponent();

            MyTextBlock.Text = message;
            FontSizeSlider.Value = StartFontSize;
            CloseButton.Focus();
        }

        private void FontSizeSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var slider = sender as Slider;
            if (slider == null) return;

            var value = slider.Value;
            MyTextBlock.FontSize = value;
        }

        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.LogFileName);
            File.WriteAllText(filename, MyTextBlock.Text);

            CloseButton_OnClick(sender, e);
            Process.Start(filename);
        }

        private void CloseButton_OnClick(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
