using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using Newtonsoft.Json;

namespace SubmissionCollector.Models
{
    public static class Extensions
    {
        public static bool IsBetween(this int number, int minimum, int maximum)
        {
            return number >= minimum && number <= maximum;
        }

        public static bool AllNull(this object[] obj)
        {
            return obj.All(item => item == null);
        }


        public static IEnumerable<int> FindAllIndexOf(this IEnumerable<string> values, string s)
        {
            return values.Select((b, i) => b.Contains(s) ? i : -1).Where(i => i != -1);
        }

        public static BitmapSource ToBitmapSource(this Bitmap bitmap)
        {
            var hBitmap = bitmap.GetHbitmap();
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }
        public static bool IsNotNaN(this double dbl)
        {

            return !double.IsNaN(dbl);
                
        }

        public static T DeepClone<T>(this T objectToClone)
        {
            var serializedJson = JsonConvert.SerializeObject(objectToClone,
                Formatting.Indented,
                new JsonSerializerSettings
                {
                    TypeNameHandling = TypeNameHandling.Objects,
                    PreserveReferencesHandling = PreserveReferencesHandling.Objects,
                    TypeNameAssemblyFormatHandling = TypeNameAssemblyFormatHandling.Simple
                });

            return JsonConvert.DeserializeObject<T>(serializedJson,
                new JsonSerializerSettings
                {
                    TypeNameHandling = TypeNameHandling.Objects,
                    PreserveReferencesHandling = PreserveReferencesHandling.Objects,
                    TypeNameAssemblyFormatHandling = TypeNameAssemblyFormatHandling.Simple
                });
        }
    }
}
