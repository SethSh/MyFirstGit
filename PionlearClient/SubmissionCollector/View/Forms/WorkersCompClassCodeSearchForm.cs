using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class WorkersCompClassCodeSearchForm : Form
    {
        public WorkersCompClassCodeSearchForm(WorkerCompClassCodeReferenceSearchDisplayer displayer)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            displayer.InitializeComponent();
            host.Child = displayer;

            displayer.CloseButton.Click += CloseButton_Click;
        }

        
        private void CloseButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
