using System.Windows;
using System.Windows.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using PionlearClient.Extensions;
using PionlearClient.KeyDataFolder;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for CedentSelector.xaml
    /// </summary>
    public partial class CedentSelector
    {
        private readonly CedentSelectorViewModel _viewModel;
        private const double IdColumnWidth = 150d;
        private const double VerticalScrollWidth = 30d;

        public CedentSelector(CedentSelectorViewModel cedentSelectorViewModel)
        {
            InitializeComponent();
            CriteriaTextBox.Text = string.Empty;

            _viewModel = cedentSelectorViewModel;
            DataContext = _viewModel;
            
            ToggleShowMyCedentsMode(_viewModel.ShowMyCedents);
            CriteriaTextBox.Focus();

            _viewModel.StatusRowLength = _viewModel.NoLength;
        }

        
        private void Commit()
        {
            if (CedentsListView.SelectedItem == null) return;

            if (_viewModel.IsSearching)
            {
                MessageHelper.Show("Cedent loading still in progress", MessageType.Stop);
                return;
            }

            var businessPartner = CedentsListView.SelectedItem as BusinessPartner;
            if (businessPartner == null) return;

            CommitSelectedItemToPackage(businessPartner);
            CommitSelectedItemToUserPreferences(businessPartner);
        }
        
        
        private static void CommitSelectedItemToPackage(BusinessPartner businessPartner)
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            package.CedentId = businessPartner.ClippedId;
            package.CedentName = businessPartner.NameAndLocation;
        }

        private void CommitSelectedItemToUserPreferences(BusinessPartner businessPartner)
        {
            var up = UserPreferences.ReadFromFile();
            up.ShowMyCedents = _viewModel.ShowMyCedents;
            
            if (up.MyCedents.Count == 0 || !up.MyCedents.Contains(businessPartner, new BusinessPartnerComparer()))
            {
                up.MyCedents.Add(businessPartner); 
            } 
            
            up.WriteToFile();
        }


        private void CedentsListView_OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            var listView = (ListView)sender;
            var gView = (GridView)listView.View;

            var workingWidth = listView.ActualWidth - VerticalScrollWidth;
            if (workingWidth < 0) return;

            var nameColumnWidth = Math.Max(workingWidth - IdColumnWidth, 0);
            gView.Columns[0].Width = IdColumnWidth;
            gView.Columns[1].Width = nameColumnWidth;
        }


        private void CedentsListView_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Commit();
        }

        private void CedentsListView_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;

            Commit();
        }

        public class CedentFinder
        {
            public static IEnumerable<BusinessPartner> Find(string cedent)
            {
                var keyDataApiWrapperClientFacade = new KeyDataApiWrapperClientFacade(ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.KeyDataBaseUrl);
                var keyDataConverter = new KeyDataConverter(keyDataApiWrapperClientFacade);
                return keyDataConverter.GetCedents(cedent);
            }
        }

        private void CommitBusinessPartnerButton_OnClick(object sender, RoutedEventArgs e)
        {
            Commit();
        }
        
        private void ShowMyCedentsCheckBox_OnClick(object sender, RoutedEventArgs e)
        {
            if (ShowMyCedentsCheckBox.IsChecked == null) return;
            ToggleShowMyCedentsMode(ShowMyCedentsCheckBox.IsChecked.Value);
        }

        private void ToggleShowMyCedentsMode(bool showMyCedents)
        {
            _viewModel.StatusRowLength = _viewModel.NoLength;
            _viewModel.CedentCount = 0;

            if (showMyCedents)
            {
                _viewModel.ShowMyCedents = true;
                _viewModel.CriteriaRowLength = _viewModel.NoLength;
                _viewModel.Criteria = string.Empty;

                var up = UserPreferences.ReadFromFile();
                up.MyCedents?.ForEach(x => x.Id = x.ClippedId);
                _viewModel.Cedents = up.MyCedents?.OrderBy(x => x.NameAndLocation);
            }
            else
            {
                _viewModel.ShowMyCedents = false;
                _viewModel.CriteriaRowLength = _viewModel.CriteriaRowLengthWhenVisible;
                _viewModel.Cedents = null;
            }
        }

        
        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            //do nothing
        }
    }

    internal class BusinessPartnerComparer : IEqualityComparer<BusinessPartner>
    {
        public bool Equals(BusinessPartner x, BusinessPartner y)
        {
            return y != null && x != null && x.Id == y.Id;
        }

        public int GetHashCode(BusinessPartner obj)
        {
            return obj.Id.GetHashCode();
        }
    }
}
