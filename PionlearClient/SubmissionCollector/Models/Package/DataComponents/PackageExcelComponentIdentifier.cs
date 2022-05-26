using System.Linq;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.Models.Package.DataComponents
{
    internal class PackageExcelComponentIdentifier
    {
        public IExcelMatrix ExcelMatrix { get; set; }
        public bool Validate(bool isQuiet)
        {
            var rangeValidator = new PackageWorksheetValidator {IsQuiet = isQuiet };
            if (!rangeValidator.Validate()) return false;
            
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            var rangeNames = package.ExcelMatrices.Select(x => x.RangeName);

            var rangeName = string.Empty;
            foreach (var item in rangeNames)
            {
                if (!item.ContainsRange(rangeValidator.SelectedRange)) continue;
                rangeName = item;
                break;
            }

            if (string.IsNullOrEmpty(rangeName))
            {
                if (!isQuiet) MessageHelper.Show(@"The selection must be within an input matrix", MessageType.Stop);
                return false;
            }

            ExcelMatrix = package.ExcelMatrices.First(x => x.RangeName == rangeName);
            return true;
        }
    }
}
