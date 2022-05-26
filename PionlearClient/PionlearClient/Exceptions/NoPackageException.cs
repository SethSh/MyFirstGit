using System;
using PionlearClient.Extensions;

namespace PionlearClient.Exceptions
{
    public class NoPackageException: Exception
    {
        private readonly long _packageId;

        public NoPackageException(long packageId)
        {
            _packageId = packageId;
        }

        public override string Message => $"{BexConstants.PackageName.ToStartOfSentence()} ID {_packageId} not found";
    }
}
