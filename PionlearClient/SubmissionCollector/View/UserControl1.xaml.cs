using System;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1
    {
        public UserControl1()
        {
            InitializeComponent();
        }
    }

    public class DemoViewModel
    {
        public string Name { get; set; }
        public DateTime TodaysDate => DateTime.Now;

        public DemoViewModel()
        {
            Name = "Seth";
        }

    }
}
