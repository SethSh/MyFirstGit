using System;
using System.Globalization;
using System.IO;
using PionlearClient;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class StackTraceLogger: WorkbookLogger
    {
        public StackTraceLogger() 
        {
            SetFileName();
            EnsureFolderExists();
        }

        private void SetFileName()
        {
            Filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexConstants.StackTraceLogFileName);
        }

        private static void EnsureFolderExists()
        {
            if (!Directory.Exists(ConfigurationHelper.AppDataFolder)) Directory.CreateDirectory(ConfigurationHelper.AppDataFolder);
        }
    }

    internal abstract class WorkbookLogger : IWorkbookLogger
    {
        protected string Filename;
        
        public void DeleteFile()
        {
            if (File.Exists(Filename)) File.Delete(Filename);
        }

        public void WriteNew(Exception ex)
        {
            DeleteFile();
            File.AppendAllLines(Filename, new[]
            {
                DateTime.Now.ToString(CultureInfo.CurrentCulture),
                string.Empty,
                ex.ToString()
            });
        }
    }

    internal interface IWorkbookLogger
    {
        void WriteNew(Exception ex);
    }
}
