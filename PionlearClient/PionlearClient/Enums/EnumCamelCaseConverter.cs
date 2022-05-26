using System;
using System.ComponentModel;

namespace PionlearClient.Enums
{
    public class EnumCamelCaseConverter : EnumConverter
    {
        public EnumCamelCaseConverter(Type type)
            : base(type)
        {
        }

        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(string))
            {
                if (value != null)
                {
                    var fi = value.GetType().GetField(value.ToString());
                    if (fi != null)
                    {
                        return SplitCamelCase(fi.Name);
                    }
                }

                return string.Empty;
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }

        public static string SplitCamelCase(string input)
        {
            return System.Text.RegularExpressions.Regex.Replace(input, "([A-Z])", " $1", System.Text.RegularExpressions.RegexOptions.Compiled).Trim();
        }
    }
}
