using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class PolicyProfileDimensionsForm : Form
    {
        public PolicyProfileDimensionsForm(PolicyProfileDimensionAlternatives alternatives)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            alternatives.InitializeComponent();
            host.Child = alternatives;

            alternatives.CloseButton.Click += CloseButton_Click;
        }

        private void CloseButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
