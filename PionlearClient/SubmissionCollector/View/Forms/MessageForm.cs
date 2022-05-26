using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class MessageForm : Form
    {
        public MessageForm(MessageBoxControl messageBoxControl)
        {
            InitializeComponent();
         
            var host = new ElementHost {Dock = DockStyle.Fill};
            Controls.Add(host);
            
            messageBoxControl.InitializeComponent();
            host.Child = messageBoxControl;

            messageBoxControl.CloseButton.Click += CloseButton_Click;
            messageBoxControl.ExportButton.Click += ExportButton_Click;
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
