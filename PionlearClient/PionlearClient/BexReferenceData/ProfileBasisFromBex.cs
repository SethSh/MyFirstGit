using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using PionlearClient.Extensions;

namespace PionlearClient.BexReferenceData
{
    public class ProfileBasisFromBex : BaseReferenceDataFromBex<ProfileUnitViewModel>
    {
        public static int DefaultCode => ReferenceData.Single(basis => basis.Name == BexConstants.PercentProfileBasisName).Id;
        public static IEnumerable<string> NamesInOrder => ReferenceData.OrderBy(x => x.DisplayOrder).Select(x => x.Name); 
        
        public ProfileBasisFromBex() : base(BexFileNames.ProfileBasisFileName)
        {
                
        }

        //ToDo remove fake profile labels call when add to bex api
        public override void GetReferenceData(string appDataFolder, string secretWord, string uwpfTokenUrl, string bexSubmissionsUrl, string bexBaseUrl)
        {
            if (ReferenceData != null) return;

            var filename = Path.Combine(appDataFolder, BexFileNames.ProfileBasisFileName);
            string json;

            try
            {
                if (File.Exists(filename) && (DateTime.Now - File.GetLastWriteTime(filename).Date).TotalDays < 30)
                {
                    json = File.ReadAllText(filename);
                    DeserializeJson(json);
                }
                else
                {
                    json = "[" +
                           "{\"id\" : 1,\"name\": \"Percent\",\"displayOrder\": 1}," +
                           "{\"id\": 2,\"name\": \"Premium\",\"displayOrder\": 2}" +
                           "]";
                    DeserializeJson(json);
                    json.WriteJsonToFile(appDataFolder, filename);
                }
            }
            catch (Exception)
            {
                json = File.ReadAllText(filename);
                DeserializeJson(json);
            }
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            throw new NotImplementedException();
        }

        public static string GetName (int basisId)
        {
            return ReferenceData.Single(x => x.Id == basisId).Name;
        }
        
    }

    public class ProfileUnitViewModel
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int DisplayOrder { get; set; }
    }
}
