using System.Windows.Forms;
using System.Windows.Forms.Integration;
using SubmissionCollector.View.Enums;

namespace SubmissionCollector.View.Forms
{
    public partial class SublineWizardForm : Form
    {
        private readonly SublineWizard _sublineWizard;
        public SublineWizardForm(SublineWizard sublineWizard)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            _sublineWizard = sublineWizard;

            sublineWizard.InitializeComponent();
            host.Child = sublineWizard;

            sublineWizard.OkButton.Click += OkButton_Click;
            sublineWizard.CancelButton.Click += CancelButton_Click;
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void OkButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (_sublineWizard.Response == FormResponse.Ok) Close();
        }
    }
}
