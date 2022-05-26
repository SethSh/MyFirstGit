using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class MessageBoxForFactorsForm : Form
    {
        public MessageBoxForFactorsForm(MessageBoxForFactors messageBoxForFactors )
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            messageBoxForFactors.InitializeComponent();
            host.Child = messageBoxForFactors;

            messageBoxForFactors.CancelButton.Click += CancelButton_Click;
            messageBoxForFactors.DeleteButton.Click += DeleteButton_Click;
            messageBoxForFactors.RenameButton.Click += RenameButton_Click;
        }

        private void RenameButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void DeleteButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
