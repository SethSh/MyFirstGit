using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class HistoryForm : Form
    {
        public HistoryForm(HistoryDisplayer historyDisplayer)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            historyDisplayer.InitializeComponent();
            host.Child = historyDisplayer;

            historyDisplayer.Close.Click += Close_Click;
            historyDisplayer.ExportButton.Click += ExportButton_Click;

        }

        private void ExportButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void Close_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
