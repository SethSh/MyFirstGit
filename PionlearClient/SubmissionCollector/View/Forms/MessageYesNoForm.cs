using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class MessageYesNoForm : Form
    {
        public MessageYesNoForm(MessageBoxYesNo messageBoxYesNo)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            messageBoxYesNo.InitializeComponent();
            host.Child = messageBoxYesNo;

            messageBoxYesNo.YesButton.Click += YesButton_Click;
            messageBoxYesNo.NoButton.Click += NoButton_Click;
        }

        private void NoButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void YesButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }
    }
}
