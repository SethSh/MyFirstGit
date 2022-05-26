using System;
using System.Linq;
using System.Reflection;

namespace SubmissionCollector.Models.DataComponents
{
    public interface IProfileFormatterHandler
    {
        int ProfileBasisId { get; }
        IProfileFormatter ProfileFormatter { get; }
        bool Handles(int profileBasisId);
    }

    public class PercentProfileFormatterHandler : IProfileFormatterHandler
    {
        public int ProfileBasisId => 1;
        public IProfileFormatter ProfileFormatter => new PercentProfileFormatter();
        public bool Handles(int profileBasisId)
        {
            return profileBasisId == ProfileBasisId;
        }
    }

    public class PremiumProfileFormatterHandler : IProfileFormatterHandler
    {
        public int ProfileBasisId => 2;
        public IProfileFormatter ProfileFormatter => new PremiumProfileFormatter();
        public bool Handles(int profileBasisId)
        {
            return profileBasisId == ProfileBasisId;
        }
    }
   
    public class ProfileFormatterFactory
    {
        public static IProfileFormatter Create(int profileBasisId)
        {
            var lookupType = typeof(IProfileFormatterHandler);
            var converters = Assembly.GetExecutingAssembly().GetTypes()
                .Where(t => lookupType.IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract).ToList();
            var handlers = converters.Select(x => (IProfileFormatterHandler)Activator.CreateInstance(x));

            var handler = handlers.SingleOrDefault(h => h.Handles(profileBasisId));
            if (handler == null) throw new ArgumentOutOfRangeException($"Can't find profile formatter with ID {profileBasisId}");

            return handler.ProfileFormatter;
        }
    }
}