using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Enums;
using SubmissionCollector.Extensions;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Package;
using SubmissionCollector.Properties;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class WorkbookUploader
    {
        private readonly IPackage _package;
        private readonly ActivityTracker _activityTracker;

        public WorkbookUploader(IPackage package)
        {
            _package = package;
            _activityTracker = new ActivityTracker();
        }

        public bool UploadSuccessful { get; set; }
        public void UploadWithProgress()
        {
            try
            {
                #region parameterize form
                var marqueeProgressBarViewModel = new MarqueeProgressBarViewModel();
                var marqueeProgressBar = new MarqueeProgressBar(marqueeProgressBarViewModel);
                var form = new MarqueeProgressForm(marqueeProgressBar)
                {
                    Text = BexConstants.ApplicationName,
                    Height = (int)FormSizeHeight.Small,
                    Width = (int)FormSizeWidth.Medium,
                    StartPosition = FormStartPosition.CenterScreen,
                    ControlBox = false,
                };
                #endregion

                var uploadSuccessful = false;
                var backgroundWorker = new BackgroundWorker { WorkerReportsProgress = true };
                backgroundWorker.DoWork += Upload;

                backgroundWorker.ProgressChanged += (sender, e) =>
                {
                    var status = e.UserState.ToString();
                    marqueeProgressBarViewModel.Status = status;
                };

                backgroundWorker.RunWorkerCompleted += (sender, e) =>
                {
                    if (e.Error == null)
                    {
                        if (!_activityTracker.EndEarly)
                        {
                            if (_activityTracker.IsAnyActivity)
                            {
                                marqueeProgressBarViewModel.Message = "Upload workbook successful";
                                marqueeProgressBarViewModel.Image = Resources.Success.ToBitmapSource();
                                uploadSuccessful = true;
                            }
                            else
                            {
                                marqueeProgressBarViewModel.Message = $"Nothing to update in the {BexConstants.ServerDatabaseName.ToLower()}";
                                marqueeProgressBarViewModel.Image = Resources.Information.ToBitmapSource();
                                uploadSuccessful = true;
                            }
                        }
                        else
                        {
                            marqueeProgressBarViewModel.Message = _activityTracker.ValidationMessage;
                            marqueeProgressBarViewModel.Image = Resources.Warning.ToBitmapSource();
                        }
                    }
                    else
                    {
                        const string startMessage = "Upload workbook failed";
                        if (e.Error == null)
                        {
                            marqueeProgressBarViewModel.Message = startMessage;
                        }
                        else
                        {
                            var message = string.Join(" ", e.Error.GetInnerExceptions().Select(x => x.Message).ToArray());
                            marqueeProgressBarViewModel.Message = $"{startMessage}{Environment.NewLine}{message}";
                        }
                        
                        marqueeProgressBarViewModel.Image = Resources.Stop.ToBitmapSource();
                    }
                    
                    form.ControlBox = true;
                    var ranSuccessfully = e.Error == null && !_activityTracker.EndEarly;
                    if (ranSuccessfully)
                    {
                        marqueeProgressBarViewModel.SetSuccessAppearance();
                        form.Height = (int)FormSizeHeight.Small;
                        Globals.Ribbons.SubmissionRibbon.SetSynchronizationButtonImage(false);
                        if (_activityTracker.IsAnyActivity)
                        {
                            _package.BexCommunications.Add(new ServerCommunication(_activityTracker.UnitsOfWork.ToString()).BexCommunicationEntry);
                        }
                    }
                    else
                    {
                        marqueeProgressBarViewModel.SetFailureAppearance();

                        var bp = MessageHelper.GetBoxProperties(marqueeProgressBarViewModel.Message);
                        form.Height = (int)bp.Height;
                        form.Width = (int)bp.Width;
                    }
                    UploadSuccessful = uploadSuccessful;
                };

                backgroundWorker.RunWorkerAsync();
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageHelper.Show($"Uploading workbook to the {BexConstants.ServerDatabaseName.ToLower()} failed: {ex.Message}", MessageType.Stop);
                UploadSuccessful = false;
            }
        }

        private void Upload(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            
            worker?.ReportProgress(0, "Checking workbook version compatibility ...");
            var unableString = $"Unable to upload workbook to the {BexConstants.ServerDatabaseName.ToLower()}: the workbook version";

            Globals.ThisWorkbook.LoadBexCompatibility();
            if (!BexCompatibility.IsConnected)
            {
                _activityTracker.ValidationMessage = $"{unableString} hasn't been verified yet.  Try again in a few seconds.";
                _activityTracker.EndEarly = true;
                return;
            }

            if (!BexCompatibility.IsCompatible)
            {
                _activityTracker.ValidationMessage = $"{unableString} {BexConstants.WorkbookVersion} is no longer valid";
                _activityTracker.EndEarly = true;
                return;
            }

            worker?.ReportProgress(0, "Gathering workbook values and validating ...");

            var packageModel = _package.CreatePackageModel();
            if (_package.ExcelValidation.Length > 0)
            {
                _activityTracker.ValidationMessage = $"{BexConstants.DataValidationTitle}\n\n{_package.ExcelValidation}";
                _activityTracker.EndEarly = true;
                return;
            }

            var validation = packageModel.Validate();
            if (validation.Length > 0)
            {
                _activityTracker.ValidationMessage = $"{BexConstants.DataValidationTitle}\n\n{validation}";
                _activityTracker.EndEarly = true;
                return;
            }
            
            worker?.ReportProgress(0, "Loading workbook content ...");
            
            var sourceManager = new BexCommunicationManager();
            sourceManager.UploadModels(packageModel, _activityTracker, worker, new StackTraceLogger());
            
        }
    }
}
