using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using PionlearClient.Extensions;

namespace SubmissionCollector.BexCommunication
{
    public class ActivityTracker
    {
        public ActivityTracker()
        {
            UnitsOfWork = new ActivityUnitsOfWork();
        }

        public string ValidationMessage { get; set; }
        public string Message { get; set; }
        public bool IsAnyActivity => UnitsOfWork.Count > 0;
        public bool EndEarly { get; set; }
        public ActivityUnitsOfWork UnitsOfWork { get; set; }
    }

    public class ActivityUnitsOfWork : ICollection<ActivityUnitOfWork>
    {
        private readonly Collection<ActivityUnitOfWork> _activityUnitsOfWork;

        public ActivityUnitsOfWork()
        {
            _activityUnitsOfWork = new Collection<ActivityUnitOfWork>();
        }

        public override string ToString()
        {
            var list = new List<string>();

            var creates = _activityUnitsOfWork.Where(unitOfWork => unitOfWork.ActivityType == ActivityType.Insert).ToList();
            if (creates.Any())
            {
                var createPackage = creates.FirstOrDefault(c => !c.SegmentId.HasValue);
                if (createPackage != null)
                {
                    list.Add($"Create {createPackage.Name}");
                }
                else
                {
                    var distinctSegmentIds = creates.Select(c => c.SegmentId).Distinct();

                    foreach (var segmentId in distinctSegmentIds)
                    {
                        var minimumPriority = creates.Where(c => c.SegmentId == segmentId).Min(c => c.Priority);
                        var units = creates.Where(c => c.SegmentId == segmentId && c.Priority == minimumPriority);
                        units.ForEach(unit => list.Add($"Create {unit.Name}"));
                    }
                }
            }

            var updates = _activityUnitsOfWork.Where(unitOfWork => unitOfWork.ActivityType == ActivityType.Update).ToList();
            if (updates.Any())
            {
                var updatePackage = updates.FirstOrDefault(c => !c.SegmentId.HasValue);
                if (updatePackage != null)
                {
                    list.Add($"Update {updatePackage.Name}");
                }
                else
                {
                    var distinctSegmentIds = updates.Select(c => c.SegmentId).Distinct();
                    foreach (var segmentId in distinctSegmentIds)
                    {
                        var minimumPriority = updates.Where(c => c.SegmentId == segmentId).Min(c => c.Priority);
                        var units = updates.Where(c => c.SegmentId == segmentId && c.Priority == minimumPriority);
                        units.ForEach(unit => list.Add($"Update {unit.Name}"));
                    }
                }
            }

            var deletes = _activityUnitsOfWork.Where(unitOfWork => unitOfWork.ActivityType == ActivityType.Delete).ToList();
            if (deletes.Any())
            {
                var deletePackage = deletes.FirstOrDefault(c => !c.SegmentId.HasValue);
                if (deletePackage != null)
                {
                    list.Add($"Delete {deletePackage.Name}");
                }
                else
                {
                    var distinctSegmentIds = deletes.Select(c => c.SegmentId).Distinct();
                    foreach (var segmentId in distinctSegmentIds)
                    {
                        var minimumPriority = deletes.Where(c => c.SegmentId == segmentId).Min(c => c.Priority);
                        var units = deletes.Where(c => c.SegmentId == segmentId && c.Priority == minimumPriority);
                        units.ForEach(unit => list.Add($"Delete {unit.Name}"));
                    }
                }
            }



            return string.Join(", ", list.ToArray());
        }

        public void AddNew(ActivityUnitOfWork unitOfWork)
        {
         if (!_activityUnitsOfWork.Contains(unitOfWork, new ActivityUnitOfWorkComparer())) _activityUnitsOfWork.Add(unitOfWork);  
        }

        public IEnumerator<ActivityUnitOfWork> GetEnumerator()
        {
            return _activityUnitsOfWork.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(ActivityUnitOfWork item)
        {
            _activityUnitsOfWork.Add(item);
        }

        public void Clear()
        {
            _activityUnitsOfWork.Clear();
        }

        public bool Contains(ActivityUnitOfWork item)
        {
            return _activityUnitsOfWork.Contains(item, new ActivityUnitOfWorkComparer());
        }

        public void CopyTo(ActivityUnitOfWork[] array, int arrayIndex)
        {
            _activityUnitsOfWork.CopyTo(array,arrayIndex);
        }

        public bool Remove(ActivityUnitOfWork item)
        {
            return _activityUnitsOfWork.Remove(item);
        }

        public int Count => _activityUnitsOfWork.Count;
        public bool IsReadOnly => false;
    }

    public class ActivityUnitOfWork
    {
        public ActivityType ActivityType { get; set; }
        public int Priority { get; set; }
        public string Name { get; set; }
        public long? SegmentId { get; set; }
    }

    public enum ActivityType
    {
        Insert,
        Update,
        Delete
    }

    internal class ActivityUnitOfWorkComparer : IEqualityComparer<ActivityUnitOfWork>
    {
        public bool Equals(ActivityUnitOfWork x, ActivityUnitOfWork y)
        {
            Debug.Assert(x != null, "x != null");
            Debug.Assert(y != null, "y != null");
            if (x.SegmentId.HasValue ^ y.SegmentId.HasValue) return false;
            if (x.SegmentId.HasValue && y.SegmentId.HasValue)
            {
                return x.ActivityType == y.ActivityType && x.Priority == y.Priority && x.Name == y.Name && x.SegmentId.Value == y.SegmentId.Value;
            }
            return x.ActivityType == y.ActivityType && x.Priority == y.Priority && x.Name == y.Name;
        }

        public int GetHashCode(ActivityUnitOfWork obj)
        {
            return $"{obj.ActivityType.ToString() + obj.Priority + obj.Name}".GetHashCode();
        }
    }
}