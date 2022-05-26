using System;

namespace PionlearClient
{
    internal static class DateExtensions
    {
        private const int MinimumAcceptableYear = 1981;
        private const int MaximumAcceptableYear = 2099;

        internal static DateTime ConvertFromDateTimeOffset(this DateTimeOffset dateTime)
        {
            if (dateTime.Offset.Equals(TimeSpan.Zero))
            {
                return dateTime.UtcDateTime;
            }

            if (dateTime.Offset.Equals(TimeZoneInfo.Local.GetUtcOffset(dateTime.DateTime)))
            {
                return DateTime.SpecifyKind(dateTime.DateTime, DateTimeKind.Local);
            }

            return dateTime.DateTime;
        }

        internal static bool IsWithinAcceptableDateRange(this DateTimeOffset? dateTime)
        {
            //ignore null, they are fine
            return !dateTime.HasValue || dateTime.Value.IsWithinAcceptableDateRange();
        }

        internal static bool IsWithinAcceptableDateRange(this DateTimeOffset dateTime)
        {
            return dateTime.Year >= MinimumAcceptableYear && dateTime.Year <= MaximumAcceptableYear;
        }
    }
}
