namespace PionlearClient.Extensions
{
    public static class StringExtensions
    {
        public static string ToStartOfSentence(this string value)
        {
            return value.Trim().Length > 0 
                ?  $"{value.Substring(0,1).ToUpper()}{value.Substring(1).ToLower()}"
                :  value;
        }

        public static string ConnectWithDash(this string s1, string s2)
        {
            return $"{s1} {BexConstants.Dash} {s2}";
        }

        
        //not using - leave in case of future use
        //public static string InsertSpacesBeforeCapitalLetters(this string item)
        //{
        //    return Regex.Replace(item, "([a-z])([A-Z])", "$1 $2");
        //}
    }
}
