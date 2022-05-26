using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for MarqueeProgressBar.xaml
    /// </summary>
    public partial class MarqueeProgressBar
    {
        private const int StartFontSize = 14;
        
        public MarqueeProgressBar(MarqueeProgressBarViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            CloseButton.Focus();

            FontSizeSlider.Value = StartFontSize;
        }
        
        private void FontSizeSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var slider = (Slider)sender;
            var value = slider.Value;
            MyTextBlock.FontSize = value;
        }

        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.LogFileName);
            File.WriteAllText(filename, MyTextBlock.Text);
            Process.Start(filename);
        }
    }
}
