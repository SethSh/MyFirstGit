using System.Text;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Enums;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class WorkbookUploaderManager
    {
        private readonly WorkbookUploader _uploader;
        
        public WorkbookUploaderManager(WorkbookUploader uploader)
        {
            _uploader = uploader;
        }

        public void Manage(IPackage package)
        {
            switch (package.AttachedToRatingAnalysisStatus)
            {
                case AttachOptions.TrueAndLocked:
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"{BexConstants.UploadName.ToStartOfSentence()}s are blocked because this {BexConstants.PackageName.ToLower()} is " +
                              $"{BexConstants.AttachName.ToLower()}ed to a locked {BexConstants.RatingAnalysisName.ToLower()}.");
                
                    MessageHelper.Show(sb.ToString(), MessageType.Warning);
                    return;
                }
                
                case AttachOptions.TrueAndUnlocked:
                {
                    var sb = new StringBuilder();
                    sb.AppendLine(package.AttachedToRatingAnalysisMessage);
                    sb.AppendLine();
                    sb.AppendLine("Are you sure you want to continue?");
                    var messageResponse = MessageHelper.ShowWithYesNo(sb.ToString());
                    if (messageResponse != DialogResult.Yes) return;
                    break;
                }
            }

            
            if (!WorkbookSaveManager.IsSavedInitially)
            {
                MessageHelper.Show($"Can't {BexConstants.UploadName.ToLower()} until this workbook is saved", MessageType.Warning);
                return;
            }

            if (WorkbookSaveManager.IsReadOnly)
            {
                MessageHelper.Show($"Can't {BexConstants.UploadName.ToLower()} a read-only workbook", MessageType.Warning);
                return;
            }

            _uploader.UploadWithProgress();
            if (_uploader.UploadSuccessful)
            {
                package.SetAllItemsToNotDirty();
            }
            else
            {
                return;
            }

            Globals.ThisWorkbook.IsCurrentlyUploading = true;
            WorkbookSaveManager.Save();
        }
    }
}