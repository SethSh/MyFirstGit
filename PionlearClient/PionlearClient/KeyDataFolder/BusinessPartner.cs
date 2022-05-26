using Newtonsoft.Json;

namespace PionlearClient.KeyDataFolder
{
    public class BusinessPartner
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public string State { get; set; }

        [JsonIgnore]
        public string NameAndLocation
        {
            get
            {
                if (string.IsNullOrWhiteSpace(City) || City == "--")
                    return Name;

                return string.IsNullOrWhiteSpace(State) ? $"{Name} {BexConstants.Dash} {City}" : $"{Name} {BexConstants.Dash} {City}, {State}";
            }
        }

        [JsonIgnore]
        public string ClippedId => Id.Substring(Id.Length - 6);
        
    }
}