using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Input;

namespace SubmissionCollector.View.Forms
{
    public partial class UnderwriterSelectorForm : Form
    {
        public UnderwriterSelectorForm(UnderwriterSelector underwriterSelector)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            underwriterSelector.InitializeComponent();
            host.Child = underwriterSelector;

            underwriterSelector.UnderwriterListView.MouseDoubleClick += UnderwriterListView_MouseDoubleClick;
            underwriterSelector.UnderwriterListView.KeyDown += UnderwriterListView_KeyDown;
            underwriterSelector.CommitUnderwriterButton.Click += CommitUnderwriterButtonClick;
            underwriterSelector.CancelButton.Click += CancelButton_Click;
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void CommitUnderwriterButtonClick(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void UnderwriterListView_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;

            Close();
        }

        private void UnderwriterListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Close();
        }
    }
}
