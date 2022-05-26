using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using PionlearClient.Extensions;

namespace SubmissionCollector.Models.Historicals
{
    public class Ledger : ILedger
    {
        private readonly IList<ILedgerItem> _ledger;

        public Ledger()
        {
            _ledger = new List<ILedgerItem>();
        }

        public void Create(int rowCount)
        {
            for (var counter = 0; counter < rowCount; counter++)
            {
                Add(LedgerItem.CreateNew(counter));
            }
        }

        public void InsertItems(int start, int rowCount)
        {
            this.Where(item => item.RowId >= start).ForEach(item => item.RowId += rowCount);
            for (var i = 0; i < rowCount; i++)
            {
                Add(LedgerItem.CreateNew(start + i));
            }

            var newLedger = this.OrderBy(item => item.RowId).ToList();

            Clear();
            newLedger.ForEach(Add);
        }

        public void RemoveItems(int start, int amount)
        {
            var newLedger = new Ledger();
            for (var i = 0; i < start; i++)
            {
                newLedger.Add(this[i]);
            }

            for (var i = start + amount; i < Count; i++)
            {
                newLedger.Add(this[i]);
            }

            newLedger.Where(x => x.RowId > start).ForEach(x => x.RowId -= amount);

            Clear();
            newLedger.ForEach(Add);
        }

        public void SetIsDirty(bool isDirty)
        {
            foreach (var item in this)
            {
                item.IsDirty = isDirty;
            }
        }

        public void SetToDirty()
        {
            SetIsDirty(true);
        }

        public void SetToNotDirty()
        {
            SetIsDirty(false);
        }

        #region list internals
        public IEnumerator<ILedgerItem> GetEnumerator()
        {
            return _ledger.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(ILedgerItem item)
        {
            _ledger.Add(item);
        }

        public void Clear()
        {
            _ledger.Clear();
        }

        public bool Contains(ILedgerItem item)
        {
            return _ledger.Contains(item);
        }

        public void CopyTo(ILedgerItem[] array, int arrayIndex)
        {
            _ledger.CopyTo(array, arrayIndex);
        }

        public bool Remove(ILedgerItem item)
        {
            return _ledger.Remove(item);
        }

        public int Count => _ledger.Count;
        public bool IsReadOnly => false;
        public int IndexOf(ILedgerItem item)
        {
            return _ledger.IndexOf(item);
        }

        public void Insert(int index, ILedgerItem item)
        {
            _ledger.Insert(index, item);
        }

        public void RemoveAt(int index)
        {
            _ledger.RemoveAt(index);
        }

        public ILedgerItem this[int index]
        {
            get => _ledger[index];
            set => _ledger[index] = value;
        }
        #endregion
    }

    public interface ILedger : IList<ILedgerItem>
    {
        void Create(int rowCount);
    }

    public class LedgerItem : ILedgerItem
    {
        public long RowId { get; set; }
        public bool IsDirty { get; set; }
        public long? SourceId { get; set; }
        public DateTime? SourceTimestamp { get; set; }
        public void Clear()
        {
            IsDirty = false;
            SourceId = new long?();
            SourceTimestamp = new DateTime?();
        }

        public static ILedgerItem CreateNew(int rowId)
        {
            return new LedgerItem {RowId = rowId, IsDirty = true};
        }
    }

    public interface ILedgerItem
    {
        bool IsDirty { get; set; }
        long RowId { get; set; }
        long? SourceId { get; set; }
        DateTime? SourceTimestamp { get; set; }

        void Clear();
    }

    internal interface IProvidesLedger
    {
        Ledger Ledger { get; set; }
    }
}
