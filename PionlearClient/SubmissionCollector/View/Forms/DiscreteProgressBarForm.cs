using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class DiscreteProgressBarForm : Form
    {
        public DiscreteProgressBarForm(DiscreteProgressBar discreteProgressBar)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            discreteProgressBar.InitializeComponent();
            host.Child = discreteProgressBar;

            discreteProgressBar.CloseButton.Click += CloseButton_Click;
        }

        private void CloseButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
