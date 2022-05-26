using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class UserPreferencesForm : Form
    {
        public UserPreferencesForm(UserPreferencesDisplayer userPreferencesDisplayer)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            userPreferencesDisplayer.InitializeComponent();
            host.Child = userPreferencesDisplayer;

            userPreferencesDisplayer.OkButton.Click += OkButton_Click;
            userPreferencesDisplayer.CancelButton.Click += CancelButton_Click;
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
