using System;
using System.Linq;
using PionlearClient.KeyDataFolder;

namespace SubmissionCollector.Models.Package
{
    public class ServerCommunication
    {
        public BexCommunicationEntry BexCommunicationEntry { get; set; }

        public ServerCommunication()
        {
            var nId = Environment.UserName;
            var underwriter = UnderwritersFromKeyData.UnderwriterReferenceData.SingleOrDefault(under => under.Code.Equals(nId));

            BexCommunicationEntry = new BexCommunicationEntry
            {
                UserName = underwriter != null ? underwriter.Name : nId,
                Timestamp = DateTime.Now
            };
        }

        public ServerCommunication(string activity) : this()
        {
            BexCommunicationEntry.Activity = activity;
        }
    }

    public class BexCommunicationEntry
    {
        public const int Padding = 25;
        public string UserName { get; set; }
        public DateTime Timestamp { get; set; }
        public string Activity { get; set; }
        public override string ToString()
        {
            var namePadded = $"{UserName}".PadRight(Padding);
            var timestampPadded = $"{Timestamp}".PadRight(Padding);
            var activityPadded = $"{Activity}".PadRight(Padding);

            return $"{namePadded}{timestampPadded}{activityPadded}";
        }
    }
}