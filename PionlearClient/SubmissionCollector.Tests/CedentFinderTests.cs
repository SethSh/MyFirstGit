using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PionlearClient.KeyDataFolder;

namespace SubmissionCollector.Tests
{
    [TestClass]
    public class CedentFinderTests
    {
        [TestMethod]
        [Ignore]
        public void TestSingleCedentId()
        {
            var keyDataApiWrapperClientFacade = new KeyDataApiWrapperClientFacade(ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.KeyDataBaseUrl);
            var keyDataConverter = new KeyDataConverter(keyDataApiWrapperClientFacade);
            var cedents = keyDataConverter.GetCedents("0000171730");

            var cedent = cedents.FirstOrDefault();
            if (cedent == null)
            {
                Assert.Fail("No business partners found");
            }

            Assert.IsTrue(cedent.Name == "RLI Insurance Company");
        }
    }
}
