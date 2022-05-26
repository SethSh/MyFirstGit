using System;
using System.Collections.ObjectModel;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ViewModel
{
    public class UmbrellaTypeAllocatorViewModel : BaseUmbrellaTypeAllocatorViewModel
    {
        public UmbrellaTypeAllocatorViewModel(ISegment segment)
        {
            Title = $"{segment.Name} {BexConstants.UmbrellaTypeName}s";

            UmbrellaItems = new ObservableCollection<UmbrellaTypeViewModelPlus>();
            foreach (var item in UmbrellaTypesFromBex.GetCommercialTypes().OrderBy(umb => umb.DisplayOrder))
            {
                UmbrellaItems.Add(new UmbrellaTypeViewModelPlus
                {
                    UmbrellaTypeCode = item.UmbrellaTypeCode,
                    UmbrellaTypeName = item.UmbrellaTypeName,
                    IsSelected = false
                });
            }

            if (!segment.UmbrellaExcelMatrix.HeaderRangeName.ExistsInWorkbook()) return;
            
            segment.UmbrellaExcelMatrix.Validate();
            foreach (var alloc in segment.UmbrellaExcelMatrix.Allocations)
            {
                var code = Convert.ToInt32(alloc.Id);

                //ignore personal
                var umbrellaItem = UmbrellaItems.SingleOrDefault(umb => umb.UmbrellaTypeCode == code);
                if (umbrellaItem != null) umbrellaItem.IsSelected = true;
            }

            OkButtonToolTip = null;
            OkButtonEnabled = true;
        }
    }

    public abstract class BaseUmbrellaTypeAllocatorViewModel : ViewModelBase, IUmbrellaSelectorViewModel
    {
        private string _title;
        private ObservableCollection<UmbrellaTypeViewModelPlus> _umbrellaItems;
        private bool _okButtonEnabled;
        private string _okButtonToolTip;

        public string Title
        {
            get => _title;
            set
            {
                _title = value;
                NotifyPropertyChanged();
            }
        }

        public ObservableCollection<UmbrellaTypeViewModelPlus> UmbrellaItems
        {
            get => _umbrellaItems;
            set
            {
                _umbrellaItems = value;
                NotifyPropertyChanged();
            }
        }

        public bool OkButtonEnabled
        {
            get => _okButtonEnabled;
            set
            {
                _okButtonEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public string OkButtonToolTip
        {
            get => _okButtonToolTip;
            set
            {
                _okButtonToolTip = value;
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("OkButtonToolTipVisibility");
            }
        }

        public bool OkButtonToolTipVisibility => !string.IsNullOrEmpty(OkButtonToolTip);
    }

    internal interface IUmbrellaSelectorViewModel
    {
        string Title { get; set; }
        ObservableCollection<UmbrellaTypeViewModelPlus> UmbrellaItems { get; set; }
    }

    public class UmbrellaTypeViewModelPlus : UmbrellaTypeViewModel
    {
        public bool IsSelected { get; set; }
    }
}
