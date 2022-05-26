using System;
using System.IO;
using static System.Configuration.ConfigurationManager;

namespace SubmissionCollector
{
    public static class ConfigurationHelper
    {
        private static string _secretWord;
        private static string _uwpfTokenUrl;
        private static string _keyDataBaseUrl;
        private static string _bexBaseUrl;
        private static string _appDataFolder;
        private static string _bexSubmissionsUrl;
        private const string BexBaseUrlName = "BexBaseUrl";
        private const string BexSubmissionsUrlName = "BexSubmissionsUrl";
        private const string UwpfTokenUrlName = "UwpfTokenUrl";
        private const string KeyDataBaseUrlName = "KeyDataBaseUrl";
        private const string SecretWordName = "SecretWord";
        private const string MunichReGroupName = "Munich Re Group";
        private const string BexName = "BEX";
        private const string VersionName = "V1";

        public static string AppDataFolder
        {
            get
            {
                if (!string.IsNullOrEmpty(_appDataFolder)) return _appDataFolder;

                var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                _appDataFolder= Path.Combine(appDataPath, MunichReGroupName, BexName, VersionName);
                return _appDataFolder;
            }
        }

        public static string BexBaseUrl
        {
            get
            {
                if (!string.IsNullOrEmpty(_bexBaseUrl)) return _bexBaseUrl;

                _bexBaseUrl = AppSettings[BexBaseUrlName];
                return _bexBaseUrl;
            }
        }

        public static string BexSubmissionsUrl
        {
            get
            {
                if (!string.IsNullOrEmpty(_bexSubmissionsUrl)) return _bexSubmissionsUrl;

                _bexSubmissionsUrl = AppSettings[BexSubmissionsUrlName];
                return _bexSubmissionsUrl;
            }
        }

        public static string KeyDataBaseUrl
        {
            get
            {
                if (!string.IsNullOrEmpty(_keyDataBaseUrl)) return _keyDataBaseUrl;

                _keyDataBaseUrl = AppSettings[KeyDataBaseUrlName];
                return _keyDataBaseUrl;
            }
        }

        public static string SecretWord
        {
            get
            {
                if (!string.IsNullOrEmpty(_secretWord)) return _secretWord;

                _secretWord = AppSettings[SecretWordName];
                return _secretWord;
            }
        }

        public static string UwpfTokenUrl
        {
            get
            {
                if (!string.IsNullOrEmpty(_uwpfTokenUrl)) return _uwpfTokenUrl;

                _uwpfTokenUrl =  AppSettings[UwpfTokenUrlName];
                return _uwpfTokenUrl;
            }
        }
    }
}
