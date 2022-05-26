using System.Windows.Forms;
using System.Windows.Forms.Integration;


namespace SubmissionCollector.View.Forms
{
    public partial class MarqueeProgressForm : Form
    {
        public MarqueeProgressForm(MarqueeProgressBar marqueeProgressBar)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            marqueeProgressBar.InitializeComponent();
            host.Child = marqueeProgressBar;

            marqueeProgressBar.CloseButton.Click += CloseButton_Click;
            marqueeProgressBar.ExportButton.Click += ExportButton_Click;
        }

        private void ExportButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void CloseButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
