using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class IsDirtyFlagsForm : Form
    {
        public IsDirtyFlagsForm(IsDirtyFlagDisplay display)
        {
            InitializeComponent();
            
            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            display.InitializeComponent();
            host.Child = display;

            display.CloseButton.Click += CloseButton_Click;
            display.ExportButton.Click += ExportButton_Click;
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
