using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PionlearClient.KeyDataFolder
{
    public class KeyDataConverter
    {
        private readonly KeyDataApiWrapperClientFacade _service;

        public KeyDataConverter(KeyDataApiWrapperClientFacade service)
        {
            _service = service;
        }

        public IEnumerable<Underwriter> GetUwpfUsers()
        {
            var keyDataUnderwriters = _service.GetAllActiveUsers();

            var underwriters = new List<Underwriter>();
            foreach (var user in keyDataUnderwriters)
            {
                var underwriter = new Underwriter
                {
                    Code = user.Key,
                    Name = Regex.Split(user.Name, " - ")[0]
                };
                underwriters.Add(underwriter);
            }
            return underwriters.OrderBy(u => u.Name);
        }

        public IEnumerable<Currency> GetCurrencies()
        {
            var currencies = _service.GetAllActiveCurrencies();
            return currencies.OrderBy(u => u.Name);
        }

        public IEnumerable<BusinessPartner> GetCedents(string nameOrKey)
        {
            var keyDataCedents = _service.GetCedents(nameOrKey);
            if (keyDataCedents == null ) return null;

            var businessPartners = new List<BusinessPartner>();
            foreach (var keyDataCedent in keyDataCedents)
            {
                var businessPartner = new BusinessPartner
                {
                    Id = keyDataCedent.Key,
                    Name = keyDataCedent.Name,
                    City = keyDataCedent.City,
                    State = keyDataCedent.State
                };
                businessPartners.Add(businessPartner);
            }
            return businessPartners.OrderBy(u => u.Name);
        }

    }


}
