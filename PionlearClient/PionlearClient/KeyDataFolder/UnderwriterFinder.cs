using System.Collections.Generic;

namespace PionlearClient.KeyDataFolder
{
    public class UnderwriterFinder
    {
        public IEnumerable<Underwriter> Find(string secretWord, string uwpfTokenUrl, string keyDataBaseUrl)
        {
            var keyDataApiWrapperClientFacade = new KeyDataApiWrapperClientFacade(secretWord, uwpfTokenUrl, keyDataBaseUrl);
            var keyDataConverter = new KeyDataConverter(keyDataApiWrapperClientFacade);
            return keyDataConverter.GetUwpfUsers();
        }
    }
}
