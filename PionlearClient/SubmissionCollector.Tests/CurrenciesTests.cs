using Microsoft.VisualStudio.TestTools.UnitTesting;
using PionlearClient;
using PionlearClient.KeyDataFolder;

namespace SubmissionCollector.Tests
{
    [TestClass]
    public class CurrenciesTests
    {
        [TestMethod]
        [Ignore]
        public void TestCurrency()
        {
            var keyDataApiWrapperClientFacade = new KeyDataApiWrapperClientFacade(ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.KeyDataBaseUrl);
            var currencies = keyDataApiWrapperClientFacade.GetAllActiveCurrencies();

            if (currencies == null)
            {
                Assert.Fail("No currencies found");
            }
            
            Assert.IsTrue(currencies != null);
        }

    }
}
