using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Input;

namespace SubmissionCollector.View.Forms
{
    public partial class CedentSelectorForm : Form
    {
        public CedentSelectorForm(CedentSelector cedentSelector)
        {
            InitializeComponent();

            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            cedentSelector.InitializeComponent();
            host.Child = cedentSelector;

            cedentSelector.CedentsListView.MouseDoubleClick += CedentsListView_MouseDoubleClick;
            cedentSelector.CedentsListView.KeyDown += CedentsListView_KeyDown;
            cedentSelector.CommitBusinessPartnerButton.Click += CommitBusinessPartnerButton_Click;
            cedentSelector.CancelButton.Click += CancelButton_Click;      
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CommitBusinessPartnerButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CedentsListView_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;

            Close();
        }

        private void CedentsListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Close();
        }
        
        
        
    }
}
