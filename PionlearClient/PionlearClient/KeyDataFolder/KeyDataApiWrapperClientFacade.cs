using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using MunichRe.KeyData.ApiWrapper;
using MunichRe.KeyData.ApiWrapper.Models;
using PionlearClient.TokenAuthentication;

// ReSharper disable RedundantNameQualifier

namespace PionlearClient.KeyDataFolder
{
    public class KeyDataApiWrapperClientFacade
    {
        private readonly IApiWrapperClient _client;
        
        public KeyDataApiWrapperClientFacade(string secretWord, string uwpfTokenUri, string keyDataUri)
        {

            _client = KeyDataApiFactory.CreateKeyApi(secretWord, uwpfTokenUri, keyDataUri);
        }

        public IEnumerable<UwpfUsersViewModel> GetAllActiveUsers()
        {
            // removed isActive: true because need SM
            var users = _client.UwpfUsers.GetAllUwpfUsers();
            return users;
        }

        public List<Currency> GetAllActiveCurrencies()
        {
            IEnumerable<CurrencyViewModel> currencyViewModels = _client.Currencies.GetAllCurrenciesTypes(isActive: true);
            var currencies = currencyViewModels.Select(item => new Currency
            {
                Name = item.Name, IsoCode = item.Key
            }).OrderBy(c => c.Name).ToList();

            return currencies;
        }

        public IEnumerable<BusinessPartnerViewModel> GetCedents(string nameOrKey)
        {
            if (int.TryParse(nameOrKey, out _))
            {
                var cedentKeyAsString = nameOrKey;
                try
                {
                    var cedent = _client.BusinessPartners.GetById(cedentKeyAsString);
                    return new[] { cedent };
                }
                catch (Microsoft.Rest.HttpOperationException ex)
                {
                    if (ex.Response.StatusCode == HttpStatusCode.NotFound)
                    {
                        return null;
                    }
                    
                    throw new Exception("Error getting cedent from key data", ex);
                }
            }

            var cedentName = $"*{nameOrKey}*";
            return _client.BusinessPartners.GetAllBusinessPartner(isActive: true, role: "Cedent", name: cedentName);
        }
       

        //public string GetJimSandor()
        //{
        //    var namePlus = _client.UwpfUsers.GetbyId("n1001471");
        //    return namePlus.Name;
        //}
    }
}
