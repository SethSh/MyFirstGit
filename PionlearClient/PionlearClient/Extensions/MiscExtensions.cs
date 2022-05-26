using System;
using System.Collections.Generic;
using System.IO;

namespace PionlearClient.Extensions
{
    public static class MiscExtensions
    {

        public static void ForEach<T>(
            this IEnumerable<T> source,
            Action<T> action)
        {
            foreach (T element in source)
                action(element);
        }
        
        public static void WriteJsonToFile(this string referenceDataAsJson, string appDataFolder, string filename)
        {
            Directory.CreateDirectory(appDataFolder);
            var fullName = Path.Combine(appDataFolder, filename);
            File.WriteAllText(fullName, referenceDataAsJson);
        }

       
        public static IEnumerable<T> Duplicates<T>(this IEnumerable<T> input)
        {
            var a = new HashSet<T>();
            var b = new HashSet<T>();
            foreach (var x in input)
            {
                if (!a.Add(x) && b.Add(x))
                    yield return x;
            }
        }
    }
}
