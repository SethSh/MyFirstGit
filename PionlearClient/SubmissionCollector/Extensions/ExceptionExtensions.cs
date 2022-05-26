using System;
using System.Collections.Generic;

namespace SubmissionCollector.Extensions
{
    public static class ExceptionExtensions
    {
        public static IEnumerable<Exception> GetInnerExceptions(this Exception ex)
        {
            var innerException = ex;
            do
            {
                yield return innerException;
                innerException = innerException.InnerException;
            } while (innerException != null);
        }

    }
}
