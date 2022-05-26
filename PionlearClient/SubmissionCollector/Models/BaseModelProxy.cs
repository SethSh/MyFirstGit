using System;
using System.Collections.Generic;
using System.ComponentModel;
using PionlearClient.Model;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.Models
{
    public abstract class BaseModelProxy: ViewModelBase, IModelProxy
    {
        private long? _sourceId;
        private DateTime? _sourceTimestamp;
        private bool _isDirty;
        private string _name;
        private long? _predecessorSourceId;

        protected BaseModelProxy()
        {
            Guid = Guid.NewGuid();
            IsDirty = true;
        }


        [Category(@"Development")]
        [Browsable(false)]
        [ReadOnly(true)]
        public virtual bool IsDirty
        {
            get => _isDirty;
            set
            {
                _isDirty = value;
                NotifyPropertyChanged();

                Globals.Ribbons.SubmissionRibbon.SetSynchronizationButtonImage(_isDirty);
            }
        }
        
        [Category(@"Development")]
        [Browsable(false)]
        [ReadOnly(true)]
        public Guid Guid { get; set; }
        
        [Category("Server")]
        [DisplayName("ID")]
        [Browsable(false)]
        public virtual long? SourceId
        {
            get => _sourceId;
            set
            {
                _sourceId = value; 
                NotifyPropertyChanged();
                // ReSharper disable once ExplicitCallerInfoArgument
                NotifyPropertyChanged("HasSourceId");
            }
        }

        [ReadOnly(true)]
        [Category("Server")]
        [DisplayName("Database Predecessor ID")]
        [Browsable(false)]
        public long? PredecessorSourceId
        {
            get => _predecessorSourceId;
            set
            {
                _predecessorSourceId = value;
                NotifyPropertyChanged();
            }
        }

        [Category("Server")]
        [DisplayName(@"Last Modified")]
        [Description("Last date and time the workbook content was sent to global system")]
        [Browsable(false)]
        public DateTime? SourceTimestamp
        {
            get => _sourceTimestamp;
            set
            {
                _sourceTimestamp = value;
                NotifyPropertyChanged();
            }
        }

        public virtual void DecoupleFromServer()
        {
            SourceId = null;
            SourceTimestamp = null;
            IsDirty = true;
        }

        [Category("Server")]
        [DisplayName(@"Coupled")]
        [Description("Coupled with global system")]
        [Browsable(true)]
        [ReadOnly(true)]
        public bool HasSourceId => SourceId.HasValue;
        
        [Category(@"Development")]
        [Browsable(false)]
        [ReadOnly(true)]
        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }
        
        public void RegisterBaseChangeEvents()
        {
            PropertyChanged += BaseProxy_PropertyChanged;
        }

        private void BaseProxy_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var list = new List<string>
            {
                "IsDirty",
                "SegmentViews",
                "IsSelected",
                "DisplayOrder",
                "IsExpanded",
                "SourceTimestamp",
                "SourceId"
            };

            if (!list.Contains(e.PropertyName))
            {
                IsDirty = true;
            }
        }
    }

    public interface IModelProxy : IModel
    {
        void DecoupleFromServer();
    }
}
