using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PionlearClient.BexReferenceData;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for MessageBoxControl.xaml
    /// </summary>
    public partial class MessageBoxControl
    {
        private const int StartFontSize = 14;

        public MessageBoxControl(IMessageBoxControlViewModel viewModel)
        {
            InitializeComponent();

            DataContext = viewModel;
            
            FakeButton.Background = new ImageBrush(viewModel.MessageType.ToResourceBitmap().ToBitmapSource());
            CloseButton.Focus();

            var vis = viewModel.ShowFontResizeAndExport ? Visibility.Visible : Visibility.Hidden;
            ExportButton.Visibility = vis;
            FontSizeSlider.Visibility = vis;
            FontSizeSlider.Value = StartFontSize;
        }
        
        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.LogFileName);
            File.WriteAllText(filename, MyTextBlock.Text);
            Process.Start(filename);
        }


        private void FontSizeSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var slider = (Slider) sender;
            var value = slider.Value;
            MyTextBlock.FontSize = value;
        }
    }
}
