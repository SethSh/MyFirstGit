using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Engine;
using SubmissionCollector.Proxy;
using Microsoft.Office.Interop;

namespace SubmissionCollector.ExcelManipulation
{
    public class ProfileWorksheetManager
    {
        public Worksheet ProfileWorksheet { get; set; }

        public ProfileWorksheetManager()
        {
            
        }

        public void CreateRiskProfileSheet(IRiskProfileProxy proxy, Worksheet worksheetSource)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;

            var visibility = worksheetSource.Visible;
            if (visibility != XlSheetVisibility.xlSheetVisible)
            {
                worksheetSource.Visible = XlSheetVisibility.xlSheetVisible;
                worksheetSource.Copy(After: Globals.ThisWorkbook.Application.ActiveSheet);
                worksheetSource.Visible = visibility;
            }
            else
            {
                worksheetSource.Copy(After: Globals.ThisWorkbook.Application.ActiveSheet);
            }
            

            ProfileWorksheet = (Worksheet)Globals.ThisWorkbook.ActiveSheet;
            Globals.ThisWorkbook.Application.Goto(ProfileWorksheet.Range["A1"]);
            ProfileWorksheet.Name = proxy.Name;

            ResetRangeNames(proxy.ShortName);

            var rangeName = "profile." + proxy.ShortName + ".header";
            var range = Globals.ThisWorkbook.Names.Item(rangeName).RefersToRange;
            var header = new string[,] { { Globals.ThisWorkbook.ThisWorkspace.SingleClientPackageProxy.Name }, { proxy.Name } };
            range.Offset[0, 1].Resize[header.Length, 1].Value = header;

            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        public void CreateRiskProfileSheet(IRiskProfileProxy proxy)
        {
            var worksheetName = RiskProfile.GetLineOfBusinessAbbreviation(proxy.LineOfBusinessName) + "Template";
            var worksheetSource = (Worksheet)Globals.ThisWorkbook.Sheets[worksheetName];
            CreateRiskProfileSheet(proxy, worksheetSource);
        }
        
        public void DuplicateRiskProfileSheet(IRiskProfileProxy proxy, IRiskProfileProxy proxySource)
        {
            proxy.WorksheetManager.CreateRiskProfileSheet(proxy, proxySource.WorksheetManager.ProfileWorksheet);
        }

        private void ResetRangeNames(string name)
        {
            if (ProfileWorksheet == null)
            {
                throw new InvalidOperationException("Can't find range names to rename");
            }

            foreach (Name item in ProfileWorksheet.Names)
            {
                var rangeNameWithSheetName = item.Name;
                var prefixEndPosition = rangeNameWithSheetName.IndexOf("!", StringComparison.Ordinal);
                var rangeName = rangeNameWithSheetName.Substring(prefixEndPosition + 1);

                var words = rangeName.Split('.');
                words[1] = name;
                var newRangeName = string.Join(".", words);

                Globals.ThisWorkbook.Names.Add(newRangeName, item.RefersToRange);
                item.Delete();
            }
        }

        public void RemoveRiskProfileSheet()
        {
            if (ProfileWorksheet == null)
            {
                throw new InvalidOperationException("Can't find profileWorksheet to delete");
            }

            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Globals.ThisWorkbook.Application.DisplayAlerts = false;

            ProfileWorksheet.Delete();
            Globals.ThisWorkbook.RemoveErrorRanges();

            Globals.ThisWorkbook.Application.DisplayAlerts = true;
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
        
    }
}
