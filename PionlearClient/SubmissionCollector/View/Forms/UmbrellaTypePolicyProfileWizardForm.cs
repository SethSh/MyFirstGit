using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class UmbrellaTypePolicyProfileWizardForm : Form
    {
        public UmbrellaTypePolicyProfileWizardForm(UmbrellaTypePolicyProfileWizard wizard)
        {
            InitializeComponent();
            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);
            
            wizard.InitializeComponent();
            host.Child = wizard;

            wizard.OkButton.Click += OkButton_Click;
            wizard.CancelButton.Click += CancelButton_Click;
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void OkButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
