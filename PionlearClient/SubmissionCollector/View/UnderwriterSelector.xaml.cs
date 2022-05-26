using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PionlearClient.KeyDataFolder;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for UnderwriterSelector.xaml
    /// </summary>
    public partial class UnderwriterSelector
    {
        private readonly UnderwriterSelectorViewModel _viewModel;
        private const double IdColumnWidth = 150d;
        private const double VerticalScrollWidth = 30d;

        public UnderwriterSelector(UnderwriterSelectorViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            _viewModel = viewModel;

            CriteriaTextBox.Focus();
        }


        private void Commit()
        {
            var underwriter = UnderwriterListView.SelectedItem as Underwriter;
            if (underwriter == null) return;

            CommitSelectedItemToPackage(underwriter);
            CommitSelectedItemToUserPreferences(underwriter);
        }

        private static void CommitSelectedItemToPackage(Underwriter underwriter)
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            
            package.ResponsibleUnderwriter = underwriter.Name;
            package.ResponsibleUnderwriterId = underwriter.Code;

        }

        private void CommitSelectedItemToUserPreferences(Underwriter underwriter)
        {
            var up = UserPreferences.ReadFromFile();
            up.ShowMyUnderwriters = _viewModel.ShowMyUnderwriters;

            if (up.MyUnderwriters.Count == 0 || !up.MyUnderwriters.Contains(underwriter, new UnderwriterComparer()))
            {
                up.MyUnderwriters.Add(underwriter);
            }

            up.WriteToFile();
        }

        private void CommitUnderwriterButton_Click(object sender, RoutedEventArgs e)
        {
            Commit();
        }

        private void UnderwriterListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Commit();
        }

        private void UnderwriterListView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;

            Commit();
        }
        
        private void UnderwriterListView_OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            var listView = (ListView)sender;
            var gView = (GridView)listView.View;

            var workingWidth = listView.ActualWidth - VerticalScrollWidth;
            if (workingWidth < 0) return;

            gView.Columns[0].Width = Math.Max(0, workingWidth - IdColumnWidth );
            gView.Columns[1].Width = IdColumnWidth;
        }

        private void ShowMyUnderwritersCheckBox_OnClick(object sender, RoutedEventArgs e)
        {
            if (ShowMyUnderwritersCheckBox.IsChecked == null) return;
            _viewModel.ShowMyUnderwriters= ShowMyUnderwritersCheckBox.IsChecked.Value;
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            //do nothing
        }
    }

    internal class UnderwriterComparer : IEqualityComparer<Underwriter>
    {
        public bool Equals(Underwriter x, Underwriter y)
        {
            return y != null && x != null && x.Code == y.Code;
        }

        public int GetHashCode(Underwriter obj)
        {
            return obj.Code.GetHashCode();
        }
    }
}
