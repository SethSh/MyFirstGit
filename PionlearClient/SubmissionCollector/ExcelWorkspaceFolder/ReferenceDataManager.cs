using System;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.BexReferenceData;
using PionlearClient.KeyDataFolder;
using SubmissionCollector.Enums;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class ReferenceDataManager
    {
        public static void Download()
        {
            Download(null);
        }

        public static void Download(BackgroundWorker worker)
        {
            const double dataCount = 13;
            var counter = 0d;
            const int sleepTime = 250;

            if (worker != null)
            {
                const string initializingMessage = "Initializing ...";
                worker.ReportProgress(Convert.ToInt16(counter++ / dataCount * 100), initializingMessage);
                Thread.Sleep(sleepTime);
            }

            var firstMessage = $"(Step 1 of 2) Refreshing reference data from {BexConstants.BexName} ...";

            var exposureBasisData = new ExposureBasisFromBex();
            exposureBasisData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var hazardCodeData = new HazardCodesFromBex();
            hazardCodeData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var historicalPeriodTypeData = new HistoricalPeriodTypesFromBex();
            historicalPeriodTypeData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var profileBasisData = new ProfileBasisFromBex();
            profileBasisData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var stateData = new StateCodesFromBex();
            stateData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl,
                ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            
            var constructionTypeData = new ConstructionTypeCodesFromBex();
            constructionTypeData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl,
                ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var occupancyTypeData = new OccupancyTypeCodesFromBex();
            occupancyTypeData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl,
                ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var protectionClassData = new ProtectionClassCodesFromBex();
            protectionClassData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl,
                ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            
            var subjectPolicyAlaeTreatmentData = new SubjectPolicyAlaeTreatmentsFromBex();
            subjectPolicyAlaeTreatmentData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var sublineCodeData = new SublineCodesFromBex();
            sublineCodeData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), firstMessage, sleepTime);

            var umbrellaTypeData = new UmbrellaTypesFromBex();
            umbrellaTypeData.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);

            
            var compClassCodesAndHazardsFromBex = new WorkersCompClassCodesAndHazardsFromBex();
            compClassCodesAndHazardsFromBex.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);

            var minnesotaRetentions = new MinnesotaRetentionsFromBex();
            minnesotaRetentions.GetReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);


            var sb = new StringBuilder();
            sb.AppendLine($"(Step 1 of 2) Refreshed reference data pre-work from {BexConstants.BexName}");
            sb.AppendLine($"(Step 2 of 2) Refreshing reference data from {BexConstants.UwpfName} {BexConstants.KeyDataName} ...");
            var secondMessage = sb.ToString();

            if (worker != null) ReportProgress(worker, Convert.ToInt16(counter++ / dataCount * 100), secondMessage, sleepTime);

            UnderwritersFromKeyData.GetUnderwriterReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.KeyDataBaseUrl);
            worker?.ReportProgress(Convert.ToInt16(counter++ / dataCount * 100), secondMessage);

            CurrenciesFromKeyData.GetCurrencyReferenceData(ConfigurationHelper.AppDataFolder, ConfigurationHelper.SecretWord,
                ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.KeyDataBaseUrl);

            sb = new StringBuilder();
            sb.AppendLine($"(Step 1 of 2) Refreshed reference data pre-work from {BexConstants.BexName}");
            sb.AppendLine($"(Step 2 of 2) Refreshed reference data from {BexConstants.UwpfName} {BexConstants.KeyDataName}");
            var thirdMessage = sb.ToString();

            worker?.ReportProgress(Convert.ToInt16(counter / dataCount * 100), thirdMessage);
        }

        public static void Delete()
        {
            // ReSharper disable once JoinDeclarationAndInitializer
            string filename;

            ExposureBasisFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.ExposureBasisFileName);
            File.Delete(filename);

            HazardCodesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.HazardCodesFileName);
            File.Delete(filename);

            StateCodesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.StatesFileName);
            File.Delete(filename);

            
            ConstructionTypeCodesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.ConstructionTypesFileName);
            File.Delete(filename);

            ProtectionClassCodesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.ProtectionClassesFileName);
            File.Delete(filename);

            OccupancyTypeCodesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.OccupancyTypesFileName);
            File.Delete(filename);


            WorkersCompClassCodesAndHazardsFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.WorkersCompClassCodesFileName);
            File.Delete(filename); 
            
            MinnesotaRetentionsFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.MinnesotaRetentionsFileName);
            File.Delete(filename);


            HistoricalPeriodTypesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.HistoricalPeriodTypesFileName);
            File.Delete(filename);

            ProfileBasisFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.ProfileBasisFileName);
            File.Delete(filename);

            
            SubjectPolicyAlaeTreatmentsFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.SubjectPolicyAlaeTreatmentsFileName);
            File.Delete(filename);

            SublineCodesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.LineSublineCodesFileName);
            File.Delete(filename);

            UmbrellaTypesFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.UmbrellaTypesFileName);
            File.Delete(filename);


            MinnesotaRetentionsFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.MinnesotaRetentionsFileName);
            File.Delete(filename);

            WorkersCompClassCodesAndHazardsFromBex.ReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.WorkersCompClassCodesFileName);
            File.Delete(filename);


            UnderwritersFromKeyData.UnderwriterReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, KeyDataConfiguration.UnderwritersFileName);
            File.Delete(filename);

            CurrenciesFromKeyData.CurrencyReferenceData = null;
            filename = Path.Combine(ConfigurationHelper.AppDataFolder, KeyDataConfiguration.CurrenciesFileName);
            File.Delete(filename);
        }

        public static void ReloadWithProgressBar(IWorkbookLogger logger)
        {
            try
            {
                ReloadWithProgressBar();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Reload reference data failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void ReloadWithProgressBar()
        {
            var b = new BackgroundWorker {WorkerReportsProgress = true};
            b.DoWork += Refresh;

            var viewModel = new DiscreteProgressBarViewModel {ButtonsRowPixels = DiscreteProgressBarViewModel.ButtonRowsPixelsCollapsed};
            var dpb = new DiscreteProgressBar(viewModel);
            var form = new DiscreteProgressBarForm(dpb)
            {
                Text = BexConstants.ApplicationName,
                Height = (int) FormSizeHeight.Small,
                Width = (int) FormSizeWidth.Medium,
                StartPosition = FormStartPosition.CenterScreen,
                ControlBox = false,
            };

            b.ProgressChanged += (sender, e) =>
            {
                viewModel.DoneAmount = e.ProgressPercentage;

                if (e.UserState == null) return;

                var message = e.UserState.ToString();
                if (!string.IsNullOrEmpty(message)) viewModel.Message = message;
            };

            b.RunWorkerCompleted += (sender, e) =>
            {
                viewModel.DoneAmount = 100;
                viewModel.ButtonsRowPixels = DiscreteProgressBarViewModel.ButtonRowsPixelsExpanded;
                form.ControlBox = true;
            };
            b.RunWorkerAsync();
            form.ShowDialog();
        }

        private static void Refresh(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;

            worker?.ReportProgress(0, new[] { "Loading reference data pre-work ...", string.Empty });
            Delete();
            Download(worker);
        }

        private static void ReportProgress(BackgroundWorker worker, int percent, string message, int sleepTime)
        {
            worker.ReportProgress(percent, message);
            Thread.Sleep(sleepTime);
        }
    }
}
