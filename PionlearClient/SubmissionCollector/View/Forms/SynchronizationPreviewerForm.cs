using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class SynchronizationPreviewerForm : Form
    {
        public SynchronizationPreviewerForm(SynchronizationPreviewer previewer)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            previewer.InitializeComponent();
            host.Child = previewer;

            previewer.OkButton.Click += OkButton_Click;
        }

        private void OkButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
