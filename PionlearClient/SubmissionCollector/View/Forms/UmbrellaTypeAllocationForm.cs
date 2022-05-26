using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace SubmissionCollector.View.Forms
{
    public partial class UmbrellaTypeAllocationForm : Form
    {
        public UmbrellaTypeAllocationForm(UmbrellaTypeAllocator typeAllocator)
        {
            InitializeComponent();
            var host = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(host);

            typeAllocator.InitializeComponent();
            host.Child = typeAllocator;

            typeAllocator.OkButton.Click += OkButton_Click;
            typeAllocator.CancelButton.Click += CancelButton_Click;
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void OkButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Close();
        }

        private void UmbrellaSelectorForm_Load(object sender, EventArgs e)
        {

        }
    }
}
