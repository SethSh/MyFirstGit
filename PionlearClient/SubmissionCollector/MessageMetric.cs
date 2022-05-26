using System.Collections.Generic;
using System.Linq;

namespace SubmissionCollector
{
    public class MessageMetric
    {
        public int LineNumber { get; set; }
        public int LineLength { get; set; }
        
        public IEnumerable<MessageMetric> GetMessageMetrics(string message)
        {
            var lines = message.Split('\n');
            var counter = 0;

            return lines.Select(line => new MessageMetric
            {
                LineNumber = counter++,
                LineLength = line.Length
            });
        }
    }
}