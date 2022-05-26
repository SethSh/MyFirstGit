using System;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    public static class DateExtensions
    {
        public static DateTimeOffset? MapToOffset(this DateTime? dateTime)
        {
            return dateTime.HasValue ? DateTime.SpecifyKind(dateTime.Value, DateTimeKind.Utc) : new DateTimeOffset?();
        }

        public static DateTimeOffset MapToOffset(this DateTime dateTime)
        {
            return DateTime.SpecifyKind(dateTime, DateTimeKind.Utc);
        }
    }
}
