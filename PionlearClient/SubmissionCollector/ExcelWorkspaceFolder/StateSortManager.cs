using System;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.StateSort;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class StateSortManager
    {
        private const string FailMessage = "State sort failed";
        
        public void SortByStateName(IWorkbookLogger logger)
        {
            try
            {
                var sorter = new StateSorterBasedOnName();
                if (!sorter.Validate()) return;

                using (new ExcelScreenUpdateDisabler())
                {
                    sorter.Sort();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show(FailMessage, MessageType.Stop);
            }
        }

        public void SortByStateNameWithCwOnTop(IWorkbookLogger logger)
        {
            try
            {
                var sorter = new StateSorterBasedOnNameWithCwOnTop();
                if (!sorter.Validate()) return;

                using (new ExcelScreenUpdateDisabler())
                {
                    sorter.Sort();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show(FailMessage, MessageType.Stop);
            }
        }

        public void SortByStateCode(IWorkbookLogger logger)
        {
            try
            {
                var sorter = new StateSorterBasedOnCode();
                if (!sorter.Validate()) return;

                using (new ExcelScreenUpdateDisabler())
                {
                    sorter.Sort();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show(FailMessage, MessageType.Stop);
            }
        }

        public void SortByStateCodeWithCwOnTop(IWorkbookLogger logger)
        {
            try
            {
                var sorter = new StateSorterBasedOnCodeWithCwOnTop();
                if (!sorter.Validate()) return;

                using (new ExcelScreenUpdateDisabler())
                {
                    sorter.Sort();
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show(FailMessage, MessageType.Stop);
            }
        }
    }
}
